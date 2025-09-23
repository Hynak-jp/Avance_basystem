'use client';

export default function LoadingScreen() {
  return (
    <div className="min-h-[60vh] grid place-items-center">
      <div className="flex flex-col items-center gap-3">
        <div className="h-8 w-8 animate-spin rounded-full border-2 border-current border-t-transparent" />
        <p className="text-sm text-muted-foreground">サインインを確認中…</p>
      </div>
    </div>
  );
}
