// src/lib/auth.ts
import type { NextAuthOptions } from 'next-auth';
import type { JWT } from 'next-auth/jwt';
import LineProvider from 'next-auth/providers/line';
import CredentialsProvider from 'next-auth/providers/credentials';

// LINEから返りうるプロファイルを型で定義
type LineProfile = {
  sub?: string; // OIDC subject
  userId?: string; // 互換
  name?: string;
  displayName?: string; // 互換
  picture?: string;
  pictureUrl?: string; // 互換
};

const STAFF_LINE_MAP: Record<string, string> = {
  'tagami@avance-jud.com': 'staff01',
  'y-itou@avance-jud.com': 'staff02',
};
function staffLineIdFromEmail(email: string): string | null {
  const key = String(email || '').trim().toLowerCase();
  return STAFF_LINE_MAP[key] || null;
}

export const authOptions: NextAuthOptions = {
  providers: [
    LineProvider({
      clientId: process.env.LINE_CLIENT_ID!,
      clientSecret: process.env.LINE_CLIENT_SECRET!,
      authorization: { params: { scope: 'openid profile' } },
      checks: ['pkce', 'state'],

      // ← any を使わず LineProfile にキャスト
      profile(profileRaw) {
        const p = profileRaw as LineProfile;
        return {
          id: p.sub ?? p.userId ?? '',
          name: p.name ?? p.displayName ?? '',
          email: null,
          image: p.picture ?? p.pictureUrl ?? null,
        };
      },
    }),
    ...(process.env.STAFF_LOGIN_ENABLED === '1'
      ? [
          CredentialsProvider({
            name: 'Staff',
            credentials: {
              email: { label: 'Email', type: 'email' },
              password: { label: 'Password', type: 'password' },
            },
            async authorize(credentials) {
              if (process.env.STAFF_LOGIN_ENABLED !== '1') return null;
              const rawEmail = String(credentials?.email || '').trim().toLowerCase();
              const rawPassword = String(credentials?.password || '');
              const whitelist = String(process.env.STAFF_LOGIN_EMAILS || '')
                .split(',')
                .map((v) => v.trim().toLowerCase())
                .filter(Boolean);
              const allow = whitelist.includes(rawEmail);
              const expectedPassword = String(process.env.STAFF_LOGIN_PASSWORD || '');
              const staffLineId = staffLineIdFromEmail(rawEmail);
              if (!allow || !expectedPassword || rawPassword !== expectedPassword) return null;
              if (!staffLineId) return null;
              return {
                id: staffLineId,
                name: rawEmail,
                email: rawEmail,
                role: 'staff',
                lineId: staffLineId,
              };
            },
          }),
        ]
      : []),
  ],

  session: { strategy: 'jwt' },

  callbacks: {
    async jwt({ token, profile, user }) {
      // token は拡張済み JWT 型として扱える
      const t = token as JWT;
      const p = profile as LineProfile | undefined;

      if (p?.sub) t.lineId = p.sub;
      if (p?.name ?? p?.displayName) t.name = p?.name ?? p?.displayName ?? null;

      const pic = p?.picture ?? p?.pictureUrl;
      if (pic) t.picture = pic;

      if (user) {
        const u = user as {
          role?: string;
          email?: string | null;
          name?: string | null;
          lineId?: string | null;
        };
        t.role = u.role ?? t.role ?? 'line';
        if (u.email) t.email = u.email;
        if (u.name) t.name = u.name;
        if (u.lineId) t.lineId = u.lineId;
      } else if (!t.role) {
        t.role = 'line';
      }

      if (t.role === 'staff' && !t.lineId && t.email) {
        const mapped = staffLineIdFromEmail(String(t.email));
        if (mapped) t.lineId = mapped;
      }

      return t;
    },

    async session({ session, token }) {
      const t = token as JWT;

      session.lineId = t.lineId ?? undefined;
      session.role = t.role as string | undefined;

      if (session.user) {
        if (t.name !== undefined) session.user.name = t.name ?? session.user.name ?? null;
        if (t.email && !session.user.email) session.user.email = t.email;
        if (!session.user.image && t.picture) session.user.image = t.picture;
      }
      return session;
    },
  },

  pages: {
    signIn: '/login',
  },
};
