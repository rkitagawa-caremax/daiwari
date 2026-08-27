export const GOOGLE_ALLOWED_ACCOUNTS = [
  { name: '後藤 祐策', email: 'y.goto@g.caremax.co.jp' },
  { name: '小倉 圭太郎', email: 'k.ogura@g.caremax.co.jp' },
  { name: '松江 浩平', email: 'k.matsue@g.caremax.co.jp' },
  { name: '大窪 嘉代', email: 'k.okubo@g.caremax.co.jp' },
  { name: '石原 佳奈', email: 'k.ishihara@g.caremax.co.jp' },
  { name: '宇原 承子', email: 's.uhara@g.caremax.co.jp' },
  { name: '尾崎 聡', email: 's.ozaki@caremax.co.jp' },
  { name: '北川 凌士', email: 'r.kitagawa@g.caremax.co.jp' }
];

export const normalizeEmail = (value) => String(value || '').trim().toLowerCase();

const GOOGLE_EMAIL_DOMAINS = ['g.caremax.co.jp', 'caremax.co.jp'];

export const expandAllowedEmailVariants = (email) => {
  const normalized = normalizeEmail(email);
  const atIndex = normalized.indexOf('@');
  if (atIndex <= 0) return [normalized];

  const localPart = normalized.slice(0, atIndex);
  const domain = normalized.slice(atIndex + 1);
  const variants = new Set([normalized]);

  if (GOOGLE_EMAIL_DOMAINS.includes(domain)) {
    GOOGLE_EMAIL_DOMAINS.forEach((nextDomain) => variants.add(`${localPart}@${nextDomain}`));
  }

  return Array.from(variants);
};

const GOOGLE_ALLOWED_EMAIL_SET = new Set(
  GOOGLE_ALLOWED_ACCOUNTS.flatMap((account) => expandAllowedEmailVariants(account.email))
);

export const isAllowedGoogleUser = (user) => {
  const email = normalizeEmail(user?.email);
  return !!email && GOOGLE_ALLOWED_EMAIL_SET.has(email);
};

export const buildGoogleAuthErrorMessage = (
  error,
  fallbackMessage = 'Googleログインに失敗しました。再度お試しください。'
) => {
  const code = String(error?.code || '').toLowerCase();
  if (code === 'auth/popup-closed-by-user' || code === 'auth/cancelled-popup-request') {
    return 'ログインがキャンセルされました。';
  }
  if (code === 'auth/unauthorized-domain') {
    const currentHost = typeof window !== 'undefined' ? window.location.hostname : '';
    return `このドメイン（${currentHost || 'unknown'}）はFirebase認証の許可ドメインに未登録です。`;
  }
  if (code === 'auth/operation-not-allowed') {
    return 'Firebase AuthenticationでGoogleログインが無効です。管理者設定を確認してください。';
  }
  if (code === 'auth/account-exists-with-different-credential') {
    return 'このメールは別の認証方式に紐づいています。管理者に連絡してください。';
  }
  if (code === 'auth/network-request-failed') {
    return 'ネットワークエラーでログインできませんでした。接続を確認して再試行してください。';
  }
  return fallbackMessage;
};
