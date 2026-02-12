'use client';
// 役割: フォーム一覧で1カードを描画し、進捗や遷移リンクの状態を可視化するクライアントコンポーネント。
// 注意: props はサーバー側で構築されるため、リンク可否や disabled 条件を変える場合は整合するサーバー処理を忘れず更新すること。
import Link from 'next/link';
import { makeProgressStore, FormStatus } from '@/lib/progressStore';
import { useUserEmailStore, isValidEmail } from '@/lib/userEmailStore';
import type { CaseFormStatus } from '@/hooks/useCaseFormsStatus';

type Props = {
  formId: string;
  title: string;
  description: string;
  baseUrl: string; // 外部フォームのURL
  lineId: string;
  caseId?: string | null;
  signedHref?: string; // 事前署名済みURL（優先）
  disabled?: boolean;
  disabledReason?: string;
  completed?: boolean;
  formKey?: string;
  storeKey?: string;
  serverStatus?: CaseFormStatus;
  draftStatus?: 'READY' | 'GENERATING' | 'ERROR' | null;
  draftViewUrl?: string;
  draftMessage?: string;
};

export default function FormCard({
  formId,
  title,
  description,
  baseUrl,
  lineId,
  caseId,
  signedHref: hrefOverride,
  disabled,
  disabledReason,
  completed,
  formKey,
  storeKey,
  serverStatus,
  draftStatus,
  draftViewUrl,
  draftMessage,
}: Props) {
  const progressKey = storeKey || formKey || formId;
  const isRepeatableDoc = String(progressKey).toLowerCase() === 'doc_payslip';
  const store = makeProgressStore(lineId)();
  const savedEmail = useUserEmailStore((state) => state.email);
  const savedOwnerLineId = useUserEmailStore((state) => state.ownerLineId);
  const status: FormStatus = store.statusByForm[progressKey] || 'not_started';
  // LIFFの仕様上、URLに改行やスペースが入るとエラーになるため、encodeURIComponentでエンコードする
  // redirectUrl は送信後に戻ってくるURL（ここでは完了ページに戻す）
  // formId も渡しておくと、完了ページでどのフォームが送信されたか分かる
  const normalizeCaseId = (value: string | null | undefined) => {
    if (!value) return '';
    const digits = String(value).replace(/\D/g, '');
    if (!digits) return '';
    return digits.slice(-4).padStart(4, '0');
  };
  const normalizedCaseId = normalizeCaseId(caseId);
  const intakeFormId = process.env.NEXT_PUBLIC_INTAKE_FORM_ID;
  const isIntakeById = Boolean(intakeFormId && formId === intakeFormId);
  const isIntakeForm =
    isIntakeById ||
    [formId, formKey, storeKey, progressKey]
      .filter((v): v is string => typeof v === 'string')
      .map((v) => v.toLowerCase())
      .some((v) => v.includes('intake')) ||
    title.includes('初回受付');
  const draftKey = String(formKey || storeKey || progressKey).trim().toLowerCase();
  const isDraftSupportedForm = new Set([
    's2002_userform',
    's2010_p1_career',
    's2010_p2_cause',
    's2011_income_m1',
    's2011_income_m2',
  ]).has(draftKey);
  // intake は常に有効扱いにするため、外部からの disabled を無視して判定する
  const internalDisabled = isIntakeForm ? false : !!disabled;
  const fallback = internalDisabled
    ? undefined
    : (() => {
        if (typeof window === 'undefined') return undefined;
        const url = new URL(baseUrl);
        url.searchParams.set('line_id[0]', lineId);
        url.searchParams.set('form_id', formId);
        if (!isIntakeForm) {
          // intake は管理画面の完了リダイレクト設定を優先するため付与しない
          const redirectUrl = new URL('/done', window.location.origin);
          redirectUrl.searchParams.set('formId', formId);
          url.searchParams.set('redirect_url[0]', redirectUrl.toString());
        }
        return url.toString();
      })();
  const signedHref = hrefOverride ?? fallback;
  const stripRedirectParam = (href: string) => {
    if (!href) return href;
    if (!/^https?:\/\//i.test(href)) return href;
    try {
      const url = new URL(href);
      url.searchParams.delete('redirect_url');
      url.searchParams.delete('redirect_url[0]');
      return url.toString();
    } catch {
      return href;
    }
  };
  const baseHref = isIntakeForm ? stripRedirectParam(signedHref ?? '') : (signedHref ?? '');
  const appendMailParam = (href: string, email: string | null) => {
    if (!href) return href;
    if (!email || !isValidEmail(email)) return href;
    if (!/^https?:\/\//i.test(href)) return href;
    try {
      const url = new URL(href);
      // FormMailer URLパラメータは「項目名[0]=初期値」がルール
      url.searchParams.delete('mail');
      url.searchParams.delete('mail[0]');
      url.searchParams.delete('メールアドレス');
      url.searchParams.delete('メールアドレス[0]');
      url.searchParams.set('メールアドレス[0]', email);
      url.searchParams.set('mail[0]', email);
      if (url.searchParams.has('case_id[0]')) {
        url.searchParams.delete('case_id');
      }
      return url.toString();
    } catch {
      return href;
    }
  };

  const overrideDone = !isRepeatableDoc && (completed ?? false);
  let effectiveStatus: FormStatus = overrideDone ? 'done' : status;

  if (serverStatus && !isRepeatableDoc) {
    if (!serverStatus.canEdit && (serverStatus.status === 'submitted' || serverStatus.status === 'closed')) {
      effectiveStatus = 'done';
    } else if (serverStatus.status === 'reopened' && effectiveStatus === 'done') {
      effectiveStatus = 'in_progress';
    }
  }

  const isDone = !isRepeatableDoc && effectiveStatus === 'done';
  const serverCanEdit =
    !isIntakeForm && !isRepeatableDoc && serverStatus ? serverStatus.canEdit : undefined;
  const baseDisabled = internalDisabled || !signedHref;
  const finalDisabled = serverCanEdit !== undefined ? !serverCanEdit || !signedHref : baseDisabled;
  const needsCaseId = !isIntakeForm;
  const hasCaseId = Boolean(normalizedCaseId);
  const caseGuardActive = needsCaseId && !hasCaseId;
  // まず通常ロジックで判定
  let isClickable = !!signedHref && !finalDisabled && !isDone && !caseGuardActive;
  // intake は href があれば必ず開けるように強制上書き
  if (isIntakeForm && signedHref && !isDone) {
    isClickable = true;
  }
  const isDisabled = !isClickable && !isDone;

  const containerClass = `rounded-lg border p-4 ${
    isClickable ? '' : 'bg-gray-50 text-gray-500 opacity-60 pointer-events-none'
  }`;
  let label: string;
  if (isDone) {
    label = '（完了）';
  } else if (caseGuardActive) {
    label = '（初回受付待ち）';
  } else if (isDisabled) {
    label = '（受付処理中）';
  } else if (effectiveStatus === 'in_progress') {
    label = '（入力中）';
  } else {
    label = '（未入力）';
  }

  if (serverStatus && !isRepeatableDoc) {
    if (!serverStatus.canEdit && (serverStatus.status === 'submitted' || serverStatus.status === 'closed')) {
      label = '（送信済み）';
    } else if (serverStatus.status === 'reopened') {
      label = '（再入力可）';
    } else if (!serverStatus.canEdit) {
      label = '（受付処理中）';
    }
  }

  const disabledMsg = (() => {
    if (caseGuardActive) {
      return 'ケースIDの登録を待っています。受付処理が完了するまでお待ちください。';
    }
    if (serverStatus && !serverStatus.canEdit && serverStatus.locked_reason) {
      return serverStatus.locked_reason;
    }
    return disabledReason || '現在は利用できません。しばらくお待ちください。';
  })();

  const reopenedHint = serverStatus?.status === 'reopened'
    ? serverStatus.reopen_until
      ? `事務所により再入力が許可されています（期限：${serverStatus.reopen_until}）`
      : '事務所により再入力が許可されています'
    : '';
  const draftHint = (() => {
    const msg = String(draftMessage || '').trim();
    if (!msg) return '';
    if (msg === 'case_not_ready') return 'ケース情報を確認中です。';
    if (msg === 'not_submitted') return '提出データを確認中です。';
    if (msg === 'draft_source_missing') return 'ドラフト元データを準備中です。';
    if (msg === 'pdf_generation_failed') return 'PDF生成に失敗しました。時間をおいて再度ご確認ください。';
    if (msg === 'status_fetch_failed') return 'ドラフト状態の取得に失敗しました。再読み込みしてください。';
    if (msg.startsWith('http_')) return 'ドラフト状態の取得中です。';
    return msg;
  })();

  const handleClick = () => {
    if (typeof window === 'undefined') return;
    if (caseGuardActive) {
      window.alert('ケースIDが未連携のため、このフォームはまだ開けません。受付処理完了後に再度お試しください。');
      return;
    }
    store.setStatus(progressKey, 'in_progress');
    const canonicalKey = formKey || storeKey || progressKey;
    const pending = {
      formKey: canonicalKey,
      storeKey: canonicalKey,
      formId,
      lineId,
      caseId: normalizedCaseId,
      savedAt: Date.now(),
    };
    try {
      window.localStorage.setItem('formlist:pendingForm', JSON.stringify(pending));
    } catch (e) {
      console.error('store pendingForm failed', e);
    }
  };

  const effectiveEmail =
    !isIntakeForm && savedOwnerLineId && savedOwnerLineId === lineId ? savedEmail : null;
  const linkHref = appendMailParam(baseHref, effectiveEmail);

  return (
    <div className={containerClass}>
      <h3 className="font-semibold">
        {title} <span className="text-sm text-gray-500">{label}</span>
      </h3>
      <p className="text-sm text-gray-600 mb-3">{description}</p>

      {isDone ? (
        <button
          className="px-3 py-1.5 bg-gray-300 text-gray-600 rounded cursor-not-allowed"
          title={overrideDone ? '受付フォームは完了しています' : '送信済みのため再送できません'}
          disabled
        >
          {overrideDone ? '受付完了しました' : '送信済み'}
        </button>
      ) : !isClickable ? (
        <>
          <button
            className="px-3 py-2 rounded bg-gray-200 cursor-not-allowed"
            title={disabledMsg}
            disabled
          >
            初回受付中
          </button>
          {disabledMsg && <div className="mt-2 text-xs text-gray-600">{disabledMsg}</div>}
        </>
      ) : (
        <Link
          href={linkHref}
          prefetch={false}
          onClick={handleClick}
          className="inline-block px-3 py-2 rounded bg-black text-white"
        >
          開く
        </Link>
      )}
      {reopenedHint && isClickable && (
        <div className="mt-2 text-xs text-gray-600">{reopenedHint}</div>
      )}
      {isDraftSupportedForm && draftStatus === 'READY' && draftViewUrl && (
        <div className="mt-3">
          <Link
            href={draftViewUrl}
            prefetch={false}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-block px-3 py-2 rounded bg-blue-700 text-white"
          >
            ドラフト閲覧（PDF）
          </Link>
        </div>
      )}
      {isDraftSupportedForm && draftStatus !== 'READY' && draftHint && (
        <div className={`mt-2 text-xs ${draftStatus === 'ERROR' ? 'text-red-700' : 'text-gray-600'}`}>
          {draftHint}
        </div>
      )}
    </div>
  );
}
