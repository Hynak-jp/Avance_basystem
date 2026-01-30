import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { notFound, redirect } from 'next/navigation';
import StaffLoginClient from './StaffLoginClient';

export default async function StaffLoginPage() {
  if (process.env.STAFF_LOGIN_ENABLED !== '1') {
    notFound();
  }
  const session = await getServerSession(authOptions);
  if (session) redirect('/form');

  return (
    <main className="container mx-auto px-4 py-10">
      <StaffLoginClient />
    </main>
  );
}
