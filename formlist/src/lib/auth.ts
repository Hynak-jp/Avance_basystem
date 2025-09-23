// src/lib/auth.ts
import type { NextAuthOptions } from 'next-auth';
import type { JWT } from 'next-auth/jwt';
import LineProvider from 'next-auth/providers/line';

// LINEから返りうるプロファイルを型で定義
type LineProfile = {
  sub?: string; // OIDC subject
  userId?: string; // 互換
  name?: string;
  displayName?: string; // 互換
  picture?: string;
  pictureUrl?: string; // 互換
};

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
  ],

  session: { strategy: 'jwt' },

  callbacks: {
    async jwt({ token, profile }) {
      // token は拡張済み JWT 型として扱える
      const t = token as JWT;
      const p = profile as LineProfile | undefined;

      if (p?.sub) t.lineId = p.sub;
      if (p?.name ?? p?.displayName) t.name = p?.name ?? p?.displayName ?? null;

      const pic = p?.picture ?? p?.pictureUrl;
      if (pic) t.picture = pic;

      return t;
    },

    async session({ session, token }) {
      const t = token as JWT;

      session.lineId = t.lineId ?? undefined;

      if (session.user) {
        if (t.name !== undefined) session.user.name = t.name ?? session.user.name ?? null;
        if (!session.user.image && t.picture) session.user.image = t.picture;
      }
      return session;
    },
  },

  pages: {
    signIn: '/login',
  },
};
