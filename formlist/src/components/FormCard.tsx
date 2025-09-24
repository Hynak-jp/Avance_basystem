'use client';

import Link from 'next/link';
import { makeProgressStore, FormStatus } from '@/lib/progressStore';

type Props = {
  formId: string;
  title: string;
  description: string;
  baseUrl: string; // 外部フォームのURL
  lineId: string;
  signedHref?: string; // 事前署名済みURL（優先）
  disabled?: boolean;
  disabledReason?: string;
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
}: Props) {
  const store = makeProgressStore(lineId)();
  const status: FormStatus = store.statusByForm[formId] || 'not_started';
  // LIFFの仕様上、URLに改行やスペースが入るとエラーになるため、encodeURIComponentでエンコードする
  // redirectUrl は送信後に戻ってくるURL（ここでは完了ページに戻す）
  // formId も渡しておくと、完了ページでどのフォームが送信されたか分かる
  const fallback = disabled
    ? undefined
    : `${baseUrl}?line_id[0]=${encodeURIComponent(lineId)}&formId=${encodeURIComponent(
        formId
      )}&redirectUrl=${encodeURIComponent('https://formlist.vercel.app/done?formId=' + formId)}`;
  const signedHref = hrefOverride ?? fallback;

  const isDone = status === 'done';
  const isClickable = !!signedHref && !disabled && !isDone;
  const isDisabled = !isClickable && !isDone;
  const containerClass = `rounded-lg border p-4 ${
    isClickable ? '' : 'bg-gray-50 text-gray-500 opacity-60 pointer-events-none'
  }`;
  const label = isDone
    ? '（送信済み）'
    : isDisabled
    ? '（受付処理中）'
    : status === 'in_progress'
    ? '（入力中）'
    : '（未入力）';
  const disabledMsg = disabledReason || '現在は利用できません。しばらくお待ちください。';

  return (
    <div className={containerClass}>
      <h3 className="font-semibold">
        {title} <span className="text-sm text-gray-500">{label}</span>
      </h3>
      <p className="text-sm text-gray-600 mb-3">{description}</p>

      {isDone ? (
        <button
          className="px-3 py-1.5 bg-gray-300 text-gray-600 rounded cursor-not-allowed"
          title="送信済みのため再送できません"
          disabled
        >
          送信済み
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
          {disabledReason && <div className="mt-2 text-xs text-gray-600">{disabledReason}</div>}
        </>
      ) : (
        <Link
          href={signedHref}
          prefetch={false}
          onClick={() => store.setStatus(formId, 'in_progress')}
          className="inline-block px-3 py-2 rounded bg-black text-white"
        >
          開く
        </Link>
      )}
    </div>
  );
}
