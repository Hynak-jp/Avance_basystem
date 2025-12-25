// 役割: 認証済みユーザー向けのフォーム一覧を SSR で描画し、署名付きリンクを提供する Next.js ページ。
// 注意: サーバーサイド fetch はヘッダ経由で PII を渡すこと・フェッチ順はロック回避のため変更に注意。
import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { redirect } from 'next/navigation';
import UserInfo from '@/components/UserInfo';
import FormProgressClient from '@/components/FormProgressClient';
import { makeFormUrl, makeIntakeUrl } from '@/lib/formUrl';
import { headers } from 'next/headers';
import { tsSec, payloadV2, signV2, b64UrlFromString } from '@/lib/sig';
import { unstable_noStore as noStore } from 'next/cache';
// headers はリダイレクト用 URL の絶対化に使用。

// サーバー動作の安定化（SSRで毎回取得）
export const dynamic = 'force-dynamic';
export const revalidate = 0;
export const fetchCache = 'force-no-store';
export const runtime = 'nodejs';

// TODO: 必要に応じて /api/forms に移行
type FormDef = {
  formId: string;
  title: string;
  description: string;
  baseUrl: string;
  formKey?: string;
  storeKey?: string;
  hidden?: boolean;
};
async function loadForms(): Promise<{ forms: FormDef[] }> {
  const forms: FormDef[] = [
    {
      formId: '302516',
      title: '初回受付フォーム',
      description: '初回受付の情報を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/47a7602b302516',
      formKey: 'intake_form',
      storeKey: 'intake_form',
    },
    {
      formId: '308335',
      title: 'S2002 破産手続開始申立書',
      description: '破産手続開始申立書の情報を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/d9b655cb308335',
      formKey: 's2002_userform',
      storeKey: 's2002_userform',
    },
    {
      formId: '315397',
      title: 'S2005 債権者一覧表',
      description: '債権者情報を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/5e2a5d6a315397',
      formKey: 's2005_creditors',
      storeKey: 's2005_creditors',
      hidden: true, // 当面は一覧に出さない（6フォーム運用）
    },
    {
      formId: '314004',
      title: 'S2006 債権者一覧表（公租公課用）',
      description: '債権者情報（公租公課用）を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/4b27d644314004',
      formKey: 's2006_creditors_public',
      storeKey: 's2006_creditors_public',
      hidden: true, // 当面は一覧に出さない（6フォーム運用）
    },
    {
      formId: '309542',
      title: 'S2010 陳述書提出フォーム(1/2)',
      description: '経歴等を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/dde2787c309542',
      formKey: 's2010_p1_career',
      storeKey: 's2010_p1_career',
    },
    {
      formId: '308949',
      title: 'S2010 陳述書提出フォーム(2/2)',
      description: '破産申立てに至った事情を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/f16a6fa1308949',
      formKey: 's2010_p2_cause',
      storeKey: 's2010_p2_cause',
    },
    {
      formId: '315503',
      title: 'S2011 家計収支提出（申立2か月前分）',
      description: '申立前2か月分の家計収支表を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/d9430156315503',
      formKey: 's2011_income_m2',
      storeKey: 's2011_income_m2',
    },
    {
      formId: '315521',
      title: 'S2011 家計収支提出（申立1か月前分）',
      description: '申立前1か月分の家計収支表を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/9e4a29d1315521',
      formKey: 's2011_income_m1',
      storeKey: 's2011_income_m1',
    },
    {
      formId: '325669',
      title: '書類提出フォーム',
      description: '給与明細などの書類をアップロードします',
      baseUrl: 'https://business.form-mailer.jp/fms/829affd7325669',
      formKey: 'supporting_documents_payslip_m1',
      storeKey: 'supporting_documents_payslip_m1',
      hidden: true, // 当面は一覧に出さない（6フォーム運用）
    },
  ];
  const visibleForms = forms.filter((form) => !form.hidden);
  return { forms: visibleForms };
}

// 署名付きの /api/status を信頼し、GAS直叩きのフォールバックは行わない

export default async function FormPage() {
  noStore();
  const { forms } = await loadForms(); // ← ここで取得
  const session = await getServerSession(authOptions);
  const lineId = session?.lineId ?? null;
  const displayName = session?.user?.name ?? '';

  if (!lineId) redirect('/login');

  // ステータス問い合わせ（caseId が無ければ intake フォームのみ表示）
  const h = await headers();
  const origin =
    process.env.NEXT_PUBLIC_BASE_URL ||
    `${h.get('x-forwarded-proto') ?? 'http'}://${h.get('host')}`;
  type StatusResponse = {
    ok?: boolean;
    caseId?: string | null;
    intakeReady?: boolean;
    caseFolderReady?: boolean;
    hasIntake?: boolean;
    activeCaseId?: string | null;
  };

  // SSR 側で 2 連打: 早すぎて反映前でも 2 発目で拾える確率を上げる
  async function callStatusOnce(lid: string, cid: string | null) {
    try {
      const ts = tsSec();
      const pld = payloadV2(lid, cid || '', ts);
      const url = new URL(`${origin}/api/status`);
      // 内部APIはヘッダで lineId/caseId を受ける方針に統一（PII のクエリ露出を避ける）
      url.searchParams.set('ts', ts);
      url.searchParams.set('sig', signV2(pld));
      url.searchParams.set('p', b64UrlFromString(pld));
      const res = await fetch(url.toString(), {
        method: 'GET',
        headers: { 'x-line-id': lid, 'x-case-id': cid || '' },
        cache: 'no-store',
        next: { revalidate: 0 },
      });
      return res.ok ? ((await res.json()) as StatusResponse) : null;
    } catch {
      return null;
    }
  }

  const primaryStatusPromise = callStatusOnce(lineId, null);
  const fallbackStatusPromise = (async () => {
    await new Promise((resolve) => setTimeout(resolve, 400));
    return callStatusOnce(lineId, null);
  })();
  let status: StatusResponse | null = null;
  try {
    status = await Promise.any([
      primaryStatusPromise.then((res) => {
        if (res) return res;
        throw new Error('empty status');
      }),
      fallbackStatusPromise.then((res) => {
        if (res) return res;
        throw new Error('empty status');
      }),
    ]);
  } catch (_) {
    const primary = await primaryStatusPromise;
    if (primary) {
      status = primary;
    } else {
      const fallback = await Promise.race([
        fallbackStatusPromise,
        new Promise<null>((resolve) => setTimeout(() => resolve(null), 800)),
      ]);
      status = fallback ?? null;
    }
  }
  if (!status?.intakeReady) {
    const candidateCaseId = status?.caseId ?? status?.activeCaseId ?? null;
    if (candidateCaseId) {
      const s2 = await callStatusOnce(lineId, candidateCaseId);
      if (s2) status = s2;
    }
  }

  const rawCaseId = status?.caseId ?? status?.activeCaseId ?? null;
  const caseId = typeof rawCaseId === 'string' && rawCaseId.length > 0 ? rawCaseId : null;
  const caseFolderReady = status?.caseFolderReady ?? Boolean(caseId);
  const intakeReady = caseFolderReady ? Boolean(status?.intakeReady) : false;
  const intakeSubmitted = status?.hasIntake ?? Boolean(caseId);
  const caseReady = Boolean(caseId);
  const intakeFormIdEnv = process.env.NEXT_PUBLIC_INTAKE_FORM_ID;
  const fallbackIntakeFormId = forms[0]?.formId;
  const preferredIntakeFormId =
    intakeFormIdEnv && intakeFormIdEnv.length > 0 ? intakeFormIdEnv : fallbackIntakeFormId;
  const intakeBase =
    process.env.NEXT_PUBLIC_INTAKE_FORM_URL ??
    forms.find((f) => f.formKey === 'intake_form')?.baseUrl ??
    'https://business.form-mailer.jp/fms/47a7602b302516';
  const userEmail = session?.user?.email ?? '';

  const formsWithHref = forms.map((f, index) => {
    const isIntakeForm =
      preferredIntakeFormId && preferredIntakeFormId.length > 0
        ? f.formId === preferredIntakeFormId
        : index === 0;
    const locked = !isIntakeForm && !caseFolderReady;
    let signedHref: string | undefined;
    const redirectUrl = new URL('/done', origin);
    redirectUrl.searchParams.set('formId', f.formId);
    if (f.formKey) redirectUrl.searchParams.set('formKey', f.formKey);
    const storeKeyParam = f.storeKey ?? f.formKey ?? f.formId;
    if (storeKeyParam) redirectUrl.searchParams.set('storeKey', storeKeyParam);
    if (caseId) redirectUrl.searchParams.set('caseId', caseId);
    redirectUrl.searchParams.set('bust', tsSec());
    if (isIntakeForm) redirectUrl.searchParams.set('form', 'intake');
    const redirectForForm = redirectUrl.toString();
    const allowPrefill = f.formKey === 's2002_userform';
    const extraPrefill =
      f.formKey === 's2002_userform' && userEmail
        ? {
            email: userEmail,
          }
        : undefined;
    if (!locked) {
      if (isIntakeForm) {
        signedHref =
          caseReady && caseId
            ? makeFormUrl(f.baseUrl, lineId!, caseId, {
                redirectUrl: redirectForForm,
                formId: f.formId,
                formKey: f.formKey,
                prefill: allowPrefill,
                extraPrefill,
                lineIdQueryKeys: [],
                caseIdQueryKeys: ['case_id[0]'],
              })
            : makeIntakeUrl(intakeBase, redirectForForm, lineId!, {
                formId: f.formId,
              });
      } else if (caseReady && caseId) {
        signedHref = makeFormUrl(f.baseUrl, lineId!, caseId, {
          redirectUrl: redirectForForm,
          formId: f.formId,
          formKey: f.formKey,
          prefill: allowPrefill,
          extraPrefill,
          lineIdQueryKeys: [],
          caseIdQueryKeys: ['case_id[0]'],
        });
      }
    }

    const intakeCompleted = isIntakeForm && caseFolderReady;
    if (intakeCompleted) {
      signedHref = undefined;
    }

    const disabled = intakeCompleted || locked || !signedHref;
    let disabledReason: string | undefined;
    if (locked) {
      disabledReason = !caseReady
        ? '受付フォームの登録が完了するまでご利用いただけません。'
        : '受付フォームを処理しています。しばらくお待ちください。';
    } else if (disabled) {
      disabledReason = '現在は利用できません。しばらくお待ちください。';
    }

    return {
      ...f,
      signedHref,
      disabled,
      disabledReason,
      completed: intakeCompleted,
      storeKey: storeKeyParam,
    };
  });

  return (
    <main className="container mx-auto px-4 py-10">
      <h1 className="text-3xl font-bold mb-6">提出フォーム一覧</h1>
      <UserInfo />
      {!caseReady ? (
        <div className="mb-6 rounded bg-blue-50 px-4 py-3 text-sm text-blue-900">
          {intakeSubmitted
            ? '受付情報を確認しています。case_id が発行されると、ほかのフォームが利用できるようになります。'
            : 'まずは「受付フォーム」をご記入ください。受付が完了すると、ほかのフォームが利用できるようになります。'}
        </div>
      ) : !caseFolderReady ? (
        <div className="mb-6 rounded bg-blue-50 px-4 py-3 text-sm text-blue-900">
          受付フォームの情報を整理しています。まもなくほかのフォームが利用できるようになります。
        </div>
      ) : !intakeReady ? (
        <div className="mb-6 rounded bg-yellow-100 px-4 py-3 text-sm text-yellow-800">
          受付フォームの処理を継続しています。ほかのフォームは利用できますが、必要に応じて再読み込みしてください。
        </div>
      ) : (
        <div className="mb-6 rounded bg-green-100 px-4 py-3 text-sm text-green-900">
          初回受付フォームの処理が完了しました。提出フォームの入力へと進んでください。
        </div>
      )}
      <FormProgressClient
        lineId={lineId!}
        displayName={displayName}
        caseId={caseId}
        forms={formsWithHref}
      />
    </main>
  );
}
