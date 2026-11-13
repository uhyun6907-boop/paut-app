import { pendingCount, signIn, useSession } from "../store/session";
import { loginBannerModel } from "../store/syncStatus";

/** 작은 '로그인 필요' 배너. 누르면 Microsoft 로그인(redirect)으로 가고, 돌아오면 대기 기록을 자동으로 보낸다. */
export function LoginBanner() {
  const session = useSession();
  const model = loginBannerModel({
    mode: session.config.syncMode,
    loginNeeded: session.loginNeeded,
    waiting: pendingCount(session.records),
    online: session.online,
  });
  if (!model.visible) return null;
  return (
    <button className="login-banner" type="button" disabled={!session.online} onClick={() => void signIn()}>
      <span className="login-banner-text">
        <strong>{model.title}</strong>
        <span>{model.detail}</span>
      </span>
      <span className="login-banner-btn">{model.button}</span>
    </button>
  );
}
