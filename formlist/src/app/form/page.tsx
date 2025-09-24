import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { redirect } from 'next/navigation';
import UserInfo from '@/components/UserInfo';
import FormProgressClient from '@/components/FormProgressClient';
import { makeFormUrl, makeIntakeUrl } from '@/lib/formUrl';
import { headers } from 'next/headers';
// headers は不要。内部 API には相対パスで十分。

// サーバー動作の安定化（SSRで毎回取得）
export const dynamic = 'force-dynamic';
export const fetchCache = 'force-no-store';
export const runtime = 'nodejs';

// TODO: 必要に応じて /api/forms に移行
type FormDef = { formId: string; title: string; description: string; baseUrl: string };
async function loadForms(): Promise<{ forms: FormDef[] }> {
  const forms: FormDef[] = [
    {
      formId: '302516',
      title: '初回受付フォーム',
      description: '初回受付の情報を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/47a7602b302516',
    },
    {
      formId: '308335',
      title: 'S2002 破産手続開始申立書',
      description: '破産手続開始申立書の情報を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/d9b655cb308335',
    },
    {
      formId: '308463',
      title: 'S2005 債権者一覧表',
      description: '債権者情報を記入します',
      baseUrl: 'https://business.form-mailer.jp/lp/47a7602b302516',
    },
    {
      formId: '309542',
      title: 'S2010 陳述書提出フォーム(1/3)',
      description: '経歴等を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/dde2787c309542',
    },
    {
      formId: '308466',
      title: 'S2010 陳述書提出フォーム(2/3)',
      description: '破産申立てに至った事情を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/f16a6fa1308949',
    },
    {
      formId: '310055',
      title: 'S2010 陳述書提出フォーム(3/3)',
      description: '免責不許可事由に関する報告を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/3bbcb828310055',
    },
    {
      formId: '308466',
      title: 'S2011 家計収支提出フォーム(1/2)',
      description: '申立前２か月分の家計収支表を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/0f10ce9b307065',
    },
    {
      formId: '308466',
      title: 'S2011 家計収支提出フォーム(2/2)',
      description: '申立前2か月分の家計収支表を記入します',
      baseUrl: 'https://business.form-mailer.jp/fms/0f10ce9b307065',
    },
    {
      formId: '307065',
      title: '書類提出フォーム',
      description: '給与明細などの書類をアップロードします',
      baseUrl: 'https://business.form-mailer.jp/fms/0f10ce9b307065',
    },
  ];
  return { forms };
}

export default async function FormPage() {
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
    hasIntake?: boolean;
    activeCaseId?: string | null;
  };

  let status: StatusResponse | null = null;
  try {
    const res = await fetch(`${origin}/api/status`, {
      method: 'GET',
      headers: { 'x-line-id': lineId },
      cache: 'no-store',
    });
    try {
      const data = (await res.json()) as StatusResponse;
      status = data;
    } catch {
      status = null;
    }
  } catch {
    status = null;
  }

  const rawCaseId = status?.caseId ?? status?.activeCaseId ?? null;
  const caseId = typeof rawCaseId === 'string' && rawCaseId.length > 0 ? rawCaseId : null;
  const intakeReady = caseId ? status?.intakeReady ?? false : false;
  const intakeSubmitted = status?.hasIntake ?? Boolean(caseId);
  const caseReady = Boolean(caseId);
  const intakeFormIdEnv = process.env.NEXT_PUBLIC_INTAKE_FORM_ID;
  const fallbackIntakeFormId = forms[0]?.formId;
  const preferredIntakeFormId =
    intakeFormIdEnv && intakeFormIdEnv.length > 0 ? intakeFormIdEnv : fallbackIntakeFormId;
  const intakeBase = process.env.NEXT_PUBLIC_INTAKE_FORM_URL!;
  const intakeRedirect = `${origin}/done?form=intake`;
  const formsWithHref = forms.map((f, index) => {
    const isIntakeForm =
      preferredIntakeFormId && preferredIntakeFormId.length > 0
        ? f.formId === preferredIntakeFormId
        : index === 0;
    let signedHref: string | undefined;
    if (isIntakeForm) {
      signedHref = caseReady
        ? makeFormUrl(f.baseUrl, lineId!, caseId!)
        : makeIntakeUrl(intakeBase, intakeRedirect);
    } else if (caseReady && intakeReady) {
      signedHref = makeFormUrl(f.baseUrl, lineId!, caseId!);
    }

    let disabled = false;
    let disabledReason: string | undefined;
    if (!isIntakeForm) {
      if (!caseReady) {
        disabled = true;
        disabledReason = '受付フォームの登録が完了するまでご利用いただけません。';
      } else if (!intakeReady) {
        disabled = true;
        disabledReason = '初回受付フォームの処理が完了するまでお待ちください。';
      }
    }
    if (!disabled && !signedHref) {
      disabled = true;
    }

    return {
      ...f,
      signedHref,
      disabled,
      disabledReason,
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
      ) : !intakeReady ? (
        <div className="mb-6 rounded bg-yellow-100 px-4 py-3 text-sm text-yellow-800">
          初回受付フォームの処理が完了するまで、ほかのフォームは操作できません。数分後に再度ご確認ください。
        </div>
      ) : null}
      <FormProgressClient lineId={lineId!} displayName={displayName} forms={formsWithHref} />
    </main>
  );
}
