'use client';

type Props = { onReset: () => void };

export default function ResetProgressButton({ onReset }: Props) {
  return (
    <button
      onClick={() => {
        if (confirm('進捗をリセットします。よろしいですか？')) onReset();
      }}
      className="mt-3 px-3 py-1.5 bg-gray-100 border rounded text-sm"
    >
      進捗をリセット（開発用）
    </button>
  );
}
