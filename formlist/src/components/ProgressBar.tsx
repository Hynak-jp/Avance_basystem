'use client';

type Props = {
  total: number;
  done: number;
};

export default function ProgressBar({ total, done }: Props) {
  const percent = total === 0 ? 0 : Math.round((done / total) * 1000) / 10;

  return (
    <div className="mb-6">
      <p className="font-semibold mb-1">
        進捗：{done}/{total} 完了
      </p>
      <div className="h-2 w-full bg-gray-200 rounded">
        <div className="h-2 bg-green-500 rounded" style={{ width: `${percent}%` }} />
      </div>
      <p className="text-sm mt-2">完了率：{percent}%</p>
      {done < total && (
        <p className="text-sm text-red-600 mt-2">
          まだ未送信のフォームがあります（{total - done}件）
        </p>
      )}
      {done === total && (
        <p className="text-sm text-green-700 mt-2">すべてのフォームが完了しました 🎉</p>
      )}
    </div>
  );
}
