import { NextRequest, NextResponse } from 'next/server';

// systemプロンプトを“日本語×数値ルール”に最適化
const SYSTEM_JA = [
  'あなたは日本の給与明細の抽出器です。必ずJSONのみ返す。未知は null。',
  '金額は円の整数。カンマ/通貨記号（¥/円）/小数点は除去し整数化する。',
  "doc_type は 'pay_slip' 固定。pay_period は 'YYYY-MM' に正規化。",
  '『総支給額(gross_amount)』『控除合計(deductions_total)』『差引支給額(net_amount)』を最優先で確定。',
  '画像と ocrText が矛盾する場合、数値や日付は ocrText を優先する。',
  'items は明細から読み取れた項目を可能な限り列挙する（例：基本給、時間外、休日、家族手当、住宅手当、健康保険、厚生年金、所得税、住民税等）。',
  '表を読み取れない場合でも、本文から拾える項目は items に入れる（見つからない項目は入れない、0円での推定は禁止）。',
].join('\n');

/** ====== Env / Defaults ====== */
const ALLOW_ORIGIN = process.env.ALLOW_ORIGIN ?? '*';
const SECRET = process.env.EXTRACT_SECRET ?? '';
// まずは安定モデル。gpt-5 を使うなら .env で OPENAI_MODEL=gpt-5 に
const MODEL = process.env.OPENAI_MODEL ?? 'gpt-4oi';

/** ====== JSON Schema（給与明細） ====== */
const paySlipSchema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  additionalProperties: false,
  required: ['doc_type', /* 'employee_name', 'employer',*/ 'pay_period', 'net_amount', 'items'],
  properties: {
    doc_type: { const: 'pay_slip' },
    source_id: { type: 'string' },
    line_id: { type: 'string' },
    employee_name: { type: ['string', 'null'] },
    employer: { type: ['string', 'null'] },
    pay_period: { type: 'string', pattern: '^\\d{4}-\\d{2}$' },
    pay_date: { type: ['string', 'null'] },
    gross_amount: { type: ['number', 'null'] },
    deductions_total: { type: ['number', 'null'] },
    net_amount: { type: 'number' },
    items: {
      type: 'array',
      minItems: 1,
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['label', 'amount'],
        properties: {
          label: { type: 'string' },
          amount: { type: 'number' }, // 控除はマイナスでもOK
        },
      },
    },
    confidence: { type: 'number', minimum: 0, maximum: 1 },
  },
} as const;

/** ====== Types ====== */
type ExtractRequestPayload = {
  imageUrl: string;
  ocrText?: string;
  lineId: string;
  sourceId: string;
};

type PayItem = { label: string; amount: number };
type PaySlip = {
  doc_type: 'pay_slip';
  source_id?: string;
  line_id?: string;
  employee_name?: string;
  employer?: string;
  pay_period?: string;
  pay_date?: string | null;
  gross_amount?: number | null;
  deductions_total?: number | null;
  net_amount?: number;
  items?: PayItem[];
  confidence?: number;
};

type ChatChoice = { message?: { content?: string } };
type ChatResponse = { choices?: ChatChoice[] };

function normalizePaySlip(json: PaySlip): PaySlip {
  // 1) doc_type を固定
  (json as PaySlip).doc_type = 'pay_slip';

  // 2) pay_period を 'YYYY- MM' に正規化（'2023年11月' 等に対応）
  const pp = json.pay_period;
  if (pp && typeof pp === 'string') {
    const m = pp.match(/(\d{4})\D+(\d{1,2})/);
    if (m) json.pay_period = `${m[1]}-${m[2].padStart(2, '0')}`;
  }

  // 3) 金額系を整数に（文字列で来た場合の保険）
  // 数値変換ヘルパーはそのまま
  const toInt = (v: unknown): number | undefined => {
    if (v == null) return undefined;
    if (typeof v === 'number' && Number.isFinite(v)) return v | 0;
    if (typeof v === 'string') {
      const s = v.replace(/[^\d-]/g, '');
      if (!s) return undefined;
      const n = parseInt(s, 10);
      return Number.isFinite(n) ? n : undefined;
    }
    return undefined;
  };

  json.gross_amount = toInt(json.gross_amount);
  json.deductions_total = toInt(json.deductions_total);
  json.net_amount = toInt(json.net_amount);

  // ← ここをガード付きに
  if (Array.isArray(json.items)) {
    json.items = json.items.map((it) => ({
      label: (it?.label ?? '').toString(),
      amount: toInt(it?.amount) ?? 0, // 必ず number
    }));
  } else {
    json.items = []; // 型が items: Array<...> | undefined なら空配列で統一
  }

  return json;
}

/** ====== CORS ====== */
function withCORS(res: NextResponse) {
  res.headers.set('Access-Control-Allow-Origin', ALLOW_ORIGIN);
  res.headers.set('Access-Control-Allow-Methods', 'POST, OPTIONS, GET');
  res.headers.set('Access-Control-Allow-Headers', 'Content-Type, X-Secret');
  res.headers.set('Access-Control-Max-Age', '86400');
  const ct = res.headers.get('Content-Type');
  if (!ct || /^application\/json\b/i.test(ct)) {
    res.headers.set('Content-Type', 'application/json; charset=utf-8');
  }
  return res;
}
export async function OPTIONS() {
  return withCORS(new NextResponse(null, { status: 200 }));
}

/** ====== Health / Diag ====== */
export async function GET() {
  return withCORS(
    NextResponse.json({
      ok: true,
      diag: {
        hasKey: !!process.env.OPENAI_API_KEY,
        model: MODEL,
        allowOrigin: ALLOW_ORIGIN,
        hasSecret: !!SECRET,
      },
    })
  );
}

/** ====== Main ====== */
export async function POST(req: NextRequest) {
  if (SECRET && req.headers.get('x-secret') !== SECRET) {
    return withCORS(new NextResponse('Unauthorized', { status: 401 }));
  }
  if (!process.env.OPENAI_API_KEY) {
    return withCORS(
      NextResponse.json({ ok: false, error: 'Missing OPENAI_API_KEY' }, { status: 500 })
    );
  }

  try {
    const { imageUrl, ocrText, lineId, sourceId } = (await req.json()) as ExtractRequestPayload;

    // (A) /file/d/.../view を直リンクへ正規化
    const imageUrlNorm = normalizeDriveUrl(imageUrl);

    // (B) HEADで Content-Type を確認（data:URLはスキップ）
    if (!imageUrlNorm.startsWith('data:')) {
      const head = await fetch(imageUrlNorm, { method: 'HEAD' });
      const ct = head.headers.get('content-type') ?? '';
      if (!/^image\/(png|jpe?g|webp|gif)\b/i.test(ct)) {
        return withCORS(
          NextResponse.json(
            {
              ok: false,
              error: `Unsupported Content-Type: ${ct}. Use a direct image URL (PNG/JPEG/WebP/GIF). Hint: https://drive.google.com/uc?export=download&id=<FILE_ID>`,
            },
            { status: 400 }
          )
        );
      }
    }

    if (!/^https?:\/\//i.test(imageUrl)) {
      return withCORS(
        NextResponse.json(
          { ok: false, error: 'imageUrl must be a public HTTP(S) URL' },
          { status: 400 }
        )
      );
    }

    // 1st: Structured Outputs（json_schema）→ 2nd: JSON mode に自動フォールバック
    const jsonText = await callChatCompletionsStrict(imageUrlNorm, ocrText).catch(async (e) => {
      console.error('[extract] schema mode failed:', e?.message ?? e);
      return await callChatCompletionsJson(imageUrlNorm, ocrText);
    });

    // Parse JSON from model
    let json: PaySlip;
    try {
      json = JSON.parse(jsonText) as PaySlip;
    } catch {
      return withCORS(
        NextResponse.json({ ok: false, error: 'Invalid JSON returned' }, { status: 500 })
      );
    }
    json = normalizePaySlip(json);

    // Fill defaults & confidence
    if (!json.doc_type) (json as PaySlip).doc_type = 'pay_slip';
    json.source_id = sourceId;
    json.line_id = lineId;

    const requiredKeys: (keyof PaySlip)[] = [
      'doc_type',
      'employee_name',
      'employer',
      'pay_period',
      'net_amount',
      'items',
    ];
    const j = json as Record<string, unknown>;
    const filledCount = requiredKeys.filter((k) => j[String(k)] != null).length;
    const fillRatio = filledCount / requiredKeys.length;

    const sumItems =
      Array.isArray(json.items) && json.items.length
        ? json.items.reduce<number>(
            (acc, it) => acc + (typeof it.amount === 'number' ? it.amount : 0),
            0
          )
        : null;

    const grossOk =
      json.gross_amount == null || sumItems == null
        ? 0.5
        : Math.abs(sumItems - (json.gross_amount ?? 0)) <= 2000
          ? 1
          : 0;

    json.confidence = Number((fillRatio * 0.7 + grossOk * 0.3).toFixed(2));

    return withCORS(NextResponse.json({ ok: true, data: json }));
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error('[extract] Fatal', msg);
    return withCORS(NextResponse.json({ ok: false, error: msg }, { status: 500 }));
  }
}

/** ====== Helpers ====== */

// 1) Chat Completions + Structured Outputs（json_schema）
async function callChatCompletionsStrict(imageUrl: string, ocrText?: string): Promise<string> {
  const body = {
    model: MODEL,
    messages: [
      { role: 'system', content: SYSTEM_JA },
      {
        role: 'user',
        content: [
          { type: 'text', text: 'この給与明細からスキーマ通りに抽出して。' },
          { type: 'image_url', image_url: { url: imageUrl, detail: 'high' } }, // ← highで精度UP
          ...(ocrText ? [{ type: 'text', text: 'ocrText:\n' + ocrText }] : []),
        ],
      },
    ],

    // ★ スキーマ厳格（まずはフォールバックしない）
    response_format: {
      type: 'json_schema',
      json_schema: { name: 'PaySlip', schema: paySlipSchema, strict: true },
    },
    temperature: 0,
  };

  const rsp = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });

  const text = await rsp.text();
  if (!rsp.ok) {
    throw new Error(text);
  }
  const data = JSON.parse(text) as ChatResponse;
  const out = data.choices?.[0]?.message?.content;
  if (!out) throw new Error('No content in chat response (schema)');
  return out;
}

// 2) フォールバック：Chat Completions の JSON モード（構造は緩いがまず通す）
async function callChatCompletionsJson(imageUrl: string, ocrText?: string): Promise<string> {
  const body = {
    model: MODEL,
    messages: [
      { role: 'system', content: SYSTEM_JA },
      {
        role: 'user',
        content: [
          { type: 'text', text: 'この給与明細からスキーマ通りに抽出して。' },
          { type: 'image_url', image_url: { url: imageUrl, detail: 'high' } }, // ← highで精度UP
          ...(ocrText ? [{ type: 'text', text: 'ocrText:\n' + ocrText }] : []),
        ],
      },
    ],
    // ★ JSON モード（多少の構造違反は許容）
    response_format: { type: 'json_object' },
    temperature: 0,
  };

  const rsp = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });

  const text = await rsp.text();
  if (!rsp.ok) {
    throw new Error(text);
  }
  const data = JSON.parse(text) as ChatResponse;
  const out = data.choices?.[0]?.message?.content;
  if (!out) throw new Error('No content in chat response (json mode)');
  return out;
}

// === Drive /view を直リンク化する（Helpers セクションに追加） ===
function normalizeDriveUrl(url: string) {
  const m = url.match(/drive\.google\.com\/file\/d\/([^/]+)/);
  return m ? `https://drive.google.com/uc?export=download&id=${m[1]}` : url;
}
