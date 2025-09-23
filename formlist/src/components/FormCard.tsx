'use client';

import Link from 'next/link';
import { makeProgressStore, FormStatus } from '@/lib/progressStore';

type Props = {
  formId: string;
  title: string;
  description: string;
  baseUrl: string; // 外部フォームのURL
  lineId: string;
  href?: string; // 事前署名済みURL（優先）
};

export default function FormCard({ formId, title, description, baseUrl, lineId, href: hrefOverride }: Props) {
  const store = makeProgressStore(lineId)();
  const status: FormStatus = store.statusByForm[formId] || 'not_started';
  // LIFFの仕様上、URLに改行やスペースが入るとエラーになるため、encodeURIComponentでエンコードする
  // redirectUrl は送信後に戻ってくるURL（ここでは完了ページに戻す）
  // formId も渡しておくと、完了ページでどのフォームが送信されたか分かる
  const fallback = `${baseUrl}?line_id[0]=${encodeURIComponent(lineId)}&formId=${encodeURIComponent(formId)}&redirectUrl=${encodeURIComponent('https://formlist.vercel.app/done?formId=' + formId)}`;
  const href = hrefOverride || fallback;

  const label =
    status === 'done' ? '（送信済み）' : status === 'in_progress' ? '（入力中）' : '（未入力）';

  return (
    <div className="border p-4 rounded">
      <h3 className="font-semibold">
        {title} <span className="text-sm text-gray-500">{label}</span>
      </h3>
      <p className="text-sm text-gray-600 mb-3">{description}</p>

      {status === 'done' ? (
        <button
          className="px-3 py-1.5 bg-gray-300 text-gray-600 rounded cursor-not-allowed"
          title="送信済みのため再送できません"
          disabled
        >
          送信済み
        </button>
      ) : (
        <Link
          href={href}
          prefetch={false}
          onClick={() => store.setStatus(formId, 'in_progress')}
          className="inline-block px-3 py-1.5 bg-blue-600 text-white rounded"
        >
          フォームへ進む
        </Link>
      )}
    </div>
  );
}
