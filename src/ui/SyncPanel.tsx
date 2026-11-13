import { retrySync, signIn, useSession } from "../store/session";
import { syncPanelModel } from "../store/syncStatus";

/** 어느 엑셀 파일이 왜 실패했는지와 '다시 시도' 버튼. 실패·오프라인 대기·명세서 시트 요청이 있으면 항상 보인다. */
export function SyncPanel() {
  const session = useSession();
  const model = syncPanelModel(session.records, session.syncReport, {
    online: session.online,
    syncing: session.syncing,
    mode: session.config.syncMode,
    loginNeeded: session.loginNeeded,
  });
  if (!model.visible) return null;
  return (
    <div className={`sync-panel ${model.tone}`} role="alert" aria-live="polite">
      <div className="sync-panel-title">{model.title}</div>
      {/* 거래명세서 그 달 시트 요청: 실패·오프라인 제목 아래에서도 따로 크게 보인다. */}
      {model.tone !== "ask" && model.prompts.map((text) => (
        <div key={text} className="sync-panel-ask">{text}</div>
      ))}
      {model.prompts.length > 0 && (
        <div className="sync-panel-hint">PC 엑셀에서 시트를 만들면 대기 기록을 다음 동기화 때 자동으로 채웁니다.</div>
      )}
      <ul className="sync-panel-lines">
        {model.lines.map((line) => (
          <li key={line.book ?? line.file} className={line.state}>
            <span className="mark">{line.state === "error" ? "✕" : line.state === "ok" ? "✓" : line.state === "waiting" ? "…" : "–"}</span>
            {line.text}
          </li>
        ))}
        {model.extra.map((text) => (
          <li key={text} className="error"><span className="mark">✕</span>{text}</li>
        ))}
      </ul>
      {model.login ? (
        <button className="btn primary block" type="button" disabled={!session.online} onClick={() => void signIn()}>다시 로그인</button>
      ) : (
      <button className="btn primary block" type="button" disabled={!model.canRetry} onClick={() => retrySync()}>
        {session.syncing ? "보내는 중…" : !session.online && session.config.syncMode === "graph" ? "연결되면 자동으로 다시 시도" : "다시 시도"}
      </button>
      )}
    </div>
  );
}
