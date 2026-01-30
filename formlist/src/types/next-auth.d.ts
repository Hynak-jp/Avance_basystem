// src/types/next-auth.d.ts
import type { DefaultSession } from 'next-auth';
import type { JWT as DefaultJWT } from 'next-auth/jwt';

declare module 'next-auth' {
  interface Session extends DefaultSession {
    lineId?: string;
    role?: string;
    user: {
      name?: string | null;
      email?: string | null;
      image?: string | null;
    };
  }
}

declare module 'next-auth/jwt' {
  interface JWT extends DefaultJWT {
    lineId?: string;
    picture?: string;
    name?: string | null;
    email?: string | null;
    role?: string;
  }
}
