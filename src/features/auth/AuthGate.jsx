import { AlertCircle } from 'lucide-react';
import { GOOGLE_ALLOWED_ACCOUNTS } from '../../config/authPolicy';

const AuthGate = ({ onGoogleSignIn, isSigningIn, errorMessage }) => {
  const handleSignIn = (event) => {
    event.preventDefault();
    onGoogleSignIn?.();
  };

  return (
    <div className="flex items-center justify-center min-h-screen relative overflow-hidden" style={{ background: 'var(--m3-surface)' }}>
      <div className="absolute top-[-15%] left-[-10%] w-[50%] h-[50%] rounded-full blur-3xl opacity-40" style={{ background: 'var(--m3-primary-container)' }} />
      <div className="absolute bottom-[-15%] right-[-10%] w-[45%] h-[45%] rounded-full blur-3xl opacity-40" style={{ background: 'var(--m3-tertiary-container)' }} />
      <div className="absolute top-[30%] right-[20%] w-[20%] h-[20%] rounded-full blur-2xl opacity-30" style={{ background: 'var(--m3-secondary-container)' }} />

      <div className="m3-card-elevated p-10 w-[420px] relative z-10 m3-animate-scale-in" style={{ borderRadius: 'var(--m3-shape-corner-xl)' }}>
        <div className="flex justify-center mb-10">
          <div className="p-1 bg-white shadow-xl" style={{ borderRadius: 'var(--m3-shape-corner-xl)' }}>
            <img src="/logo.jpg" alt="台割君" className="w-40 h-40 object-contain" style={{ borderRadius: 'calc(var(--m3-shape-corner-xl) - 4px)' }} draggable={false} />
          </div>
        </div>
        <p className="text-base mb-6 text-center font-medium" style={{ color: 'var(--m3-on-surface-variant)' }}>許可されたGoogleアカウントでログインしてください</p>

        <form onSubmit={handleSignIn} className="space-y-6">
          <div className="rounded-lg border px-4 py-3" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
            <p className="text-xs font-bold mb-2" style={{ color: 'var(--m3-on-surface-variant)' }}>許可ユーザー</p>
            <ul className="space-y-1">
              {GOOGLE_ALLOWED_ACCOUNTS.map((account) => (
                <li key={account.email} className="text-[11px] font-medium" style={{ color: 'var(--m3-on-surface)' }}>
                  <span className="font-medium" style={{ color: 'var(--m3-on-surface)' }}>{account.name}</span>
                </li>
              ))}
            </ul>
          </div>

          {errorMessage && (
            <div className="flex items-center justify-center gap-3 text-sm p-4 m3-animate-fade-in" style={{ background: 'var(--m3-error-container)', color: 'var(--m3-on-error-container)', borderRadius: 'var(--m3-shape-corner-md)' }}>
              <AlertCircle size={18} />
              <span className="font-medium">{errorMessage}</span>
            </div>
          )}
          <button
            type="submit"
            disabled={isSigningIn}
            className="w-full p-4 font-medium text-base transition-all duration-200 hover:shadow-lg active:scale-[0.98]"
            style={{
              background: 'var(--m3-primary)',
              color: 'var(--m3-on-primary)',
              borderRadius: 'var(--m3-shape-corner-full)',
              boxShadow: 'var(--m3-elevation-2)'
            }}
          >
            {isSigningIn ? 'Googleログイン中...' : 'Googleでログイン'}
          </button>
          <p className="text-[11px] text-center" style={{ color: 'var(--m3-on-surface-variant)' }}>
            一度ログインすると次回以降は自動ログインされます
          </p>
        </form>
      </div>
    </div>
  );
};

export default AuthGate;
