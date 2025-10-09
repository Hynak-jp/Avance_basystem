'use client';

import Link from 'next/link';
import { makeProgressStore, FormStatus } from '@/lib/progressStore';
import type { CaseFormStatus } from '@/hooks/useCaseFormsStatus';

type Props = {
  formId: string;
  title: string;
  description: string;
  baseUrl: string; // 外部フォームのURL
  lineId: string;
  signedHref?: string; // 事前署名済みURL（優先）
  disabled?: boolean;
  disabledReason?: string;
  completed?: boolean;
  formKey?: string;
  storeKey?: string;
  serverStatus?: CaseFormStatus;
};

export default function FormCard({
  formId,
  title,
  description,
  baseUrl,
  lineId,
  signedHref: hrefOverride,
  disabled,
  disabledReason,
  completed,
  formKey,
  storeKey,
  serverStatus,
}: Props) {
  const progressKey = storeKey || formKey || formId;
  const store = makeProgressStore(lineId)();
  const status: FormStatus = store.statusByForm[progressKey] || 'not_started';
  // LIFFの仕様上、URLに改行やスペースが入るとエラーになるため、encodeURIComponentでエンコードする
  // redirectUrl は送信後に戻ってくるURL（ここでは完了ページに戻す）
  // formId も渡しておくと、完了ページでどのフォームが送信されたか分かる
  const fallback = disabled
    ? undefined
    : (() => {
        const url = new URL(baseUrl);
        const redirectUrl = new URL('https://formlist.vercel.app/done');
        redirectUrl.searchParams.set('formId', formId);
        if (formKey) redirectUrl.searchParams.set('formKey', formKey);
        url.searchParams.set('line_id[0]', lineId);
        url.searchParams.set('form_id', formId);
        url.searchParams.set('redirect_url[0]', redirectUrl.toString());
        return url.toString();
      })();
  const signedHref = hrefOverride ?? fallback;

  const overrideDone = completed ?? false;
  let effectiveStatus: FormStatus = overrideDone ? 'done' : status;

  if (serverStatus) {
    if (!serverStatus.canEdit && (serverStatus.status === 'submitted' || serverStatus.status === 'closed')) {
      effectiveStatus = 'done';
    } else if (serverStatus.status === 'reopened' && effectiveStatus === 'done') {
      effectiveStatus = 'in_progress';
    }
  }

  const isDone = effectiveStatus === 'done';
  const serverCanEdit = serverStatus ? serverStatus.canEdit : undefined;
  const baseDisabled = disabled || !signedHref;
  const finalDisabled = serverCanEdit !== undefined ? !serverCanEdit || !signedHref : baseDisabled;
  const isClickable = !!signedHref && !finalDisabled && !isDone;
  const isDisabled = !isClickable && !isDone;

  const containerClass = `rounded-lg border p-4 ${
    isClickable ? '' : 'bg-gray-50 text-gray-500 opacity-60 pointer-events-none'
  }`;
  let label = isDone
    ? '（完了）'
    : isDisabled
    ? '（受付処理中）'
    : effectiveStatus === 'in_progress'
    ? '（入力中）'
    : '（未入力）';

  if (serverStatus) {
    if (!serverStatus.canEdit && (serverStatus.status === 'submitted' || serverStatus.status === 'closed')) {
      label = '（送信済み）';
    } else if (serverStatus.status === 'reopened') {
      label = '（再入力可）';
    } else if (!serverStatus.canEdit) {
      label = '（受付処理中）';
    }
  }

  const disabledMsg = (() => {
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
          href={signedHref}
          prefetch={false}
          onClick={() => store.setStatus(progressKey, 'in_progress')}
          className="inline-block px-3 py-2 rounded bg-black text-white"
        >
          開く
        </Link>
      )}
      {reopenedHint && isClickable && (
        <div className="mt-2 text-xs text-gray-600">{reopenedHint}</div>
      )}
    </div>
  );
}
