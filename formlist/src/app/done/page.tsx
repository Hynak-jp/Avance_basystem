import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import DoneClient from './DoneClient';
import { redirect } from 'next/navigation';

export default async function DonePage() {
  const session = await getServerSession(authOptions);
  const lineId = session?.lineId as string | undefined;
  if (!lineId) redirect('/login');
  return <DoneClient lineId={lineId} />;
}
