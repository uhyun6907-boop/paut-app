import { useEffect, useRef, useState } from "react";
import { bindOnline, dismissToast, initSession, removeRecord, resetSamples, syncQueue, useSession } from "../store/session";
import { headerStatus } from "../store/syncStatus";
import { ExcelScreen } from "./ExcelScreen";
import { EntryView } from "./EntryView";
import { RefScreen } from "./RefScreen";
import { ListScreen } from "./ListScreen";
import { SettingsScreen } from "./SettingsScreen";
import { LoginBanner } from "./LoginBanner";
import { SyncPanel } from "./SyncPanel";
import { useDrill } from "./drill";

type Tab = "entry" | "list" | "log" | "bill" | "settings" | "ref";

export function App() {
  const session = useSession();
  const [tab, setTab] = useState<Tab>("entry");
  const backTab = useRef<Tab>("entry");
  if (tab !== "settings") backTab.current = tab;
  useDrill(tab === "settings", () => setTab(backTab.current));

  useEffect(() => {
    void initSession();
    return bindOnline();
  }, []);

  useEffect(() => {
    const typing = (event: Event) => (event.target as HTMLElement | null)?.closest("input, textarea, select, [contenteditable='true']");
    const stop = (event: Event) => {
      if (typing(event)) return;
      event.preventDefault();
    };
    const mark = (event: Event) => {
      const el = (event.target as HTMLElement | null)?.closest("button, a, summary");
      if (el) el.setAttribute("draggable", "false");
    };
    document.addEventListener("contextmenu", stop);
    document.addEventListener("selectstart", stop);
    document.addEventListener("dragstart", stop);
    document.addEventListener("pointerdown", mark, true);
    return () => {
      document.removeEventListener("contextmenu", stop);
      document.removeEventListener("selectstart", stop);
      document.removeEventListener("dragstart", stop);
      document.removeEventListener("pointerdown", mark, true);
    };
  }, []);

  useEffect(() => {
    if (!session.toast) return;
    const timer = window.setTimeout(() => dismissToast(), session.toast.undoId ? 6000 : 1600);
    return () => window.clearTimeout(timer);
  }, [session.toast]);

  const status = headerStatus({
    mode: session.config.syncMode,
    online: session.online,
    syncing: session.syncing,
    loginNeeded: session.loginNeeded,
    account: session.account,
    queue: syncQueue(session.records),
    auth: session.auth,
  });
  const dot = status.tone;
  const syncText = status.text;

  if (!session.ready) {
    return (
      <div className="app">
        <div className="loading">
          <h1>오늘 검사</h1>
          <p className="nav-hint">현장 기록을 읽고 있습니다…</p>
        </div>
      </div>
    );
  }

  if (session.error) {
    return (
      <div className="app">
        <div className="loading">
          <h1>기록을 열지 못했습니다</h1>
          <p className="err">{session.error}</p>
          <button className="btn primary" type="button" onClick={() => void resetSamples().then(() => window.location.reload())}>샘플로 다시 시작</button>
        </div>
      </div>
    );
  }

  return (
    <div className="app">
      <header className={`ios-head${tab === "entry" ? " slim" : ""}`}>
        {tab === "entry" ? (
          <>
            <div className="sync-line"><span className={`sync-dot ${dot}`} /> {syncText}</div>
            <button className="gear" type="button" aria-label="설정" onClick={() => setTab("settings")}>⚙</button>
          </>
        ) : (
          <>
            <div className="ios-title-row">
              <h1>{tab === "list" ? "리스트" : tab === "log" ? "일지" : tab === "bill" ? "명세서" : tab === "settings" ? "설정" : "참고"}</h1>
              <button className="gear" type="button" aria-label="설정" onClick={() => setTab(tab === "settings" ? "entry" : "settings")}>⚙</button>
            </div>
            <div className="sync-line"><span className={`sync-dot ${dot}`} /> {syncText}</div>
          </>
        )}
      </header>
      <LoginBanner />
      <SyncPanel />
      <div className="panel active">
        {tab === "entry" && <EntryView />}
        {tab === "list" && <ListScreen />}
        {tab === "log" && <ExcelScreen part="log" />}
        {tab === "bill" && <ExcelScreen part="bill" />}
        {tab === "ref" && <RefScreen />}
        {tab === "settings" && (
          <>
            <SettingsScreen key={session.config.clientId + session.account} />
            <button className="btn block" type="button" onClick={() => setTab("ref")}>관경·스캔 참고</button>
            <button className="btn block" type="button" onClick={() => setTab("log")}>작업일지 미리보기</button>
          </>
        )}
      </div>
      {tab !== "settings" && (
        <nav className="tabbar">
          <button type="button" className={tab === "entry" ? "on" : ""} onClick={() => setTab("entry")}><span>▦</span>입력</button>
          <button type="button" className={tab === "ref" ? "on" : ""} onClick={() => setTab("ref")}><span>Ø</span>참고</button>
          <button type="button" className={tab === "list" ? "on" : ""} onClick={() => setTab("list")}><span>▤</span>리스트</button>
          <button type="button" className={tab === "bill" ? "on" : ""} onClick={() => setTab("bill")}><span>₩</span>명세서</button>
        </nav>
      )}
      {session.toast && (
        <button
          className="toast"
          type="button"
          onClick={() => {
            if (session.toast?.undoId) void removeRecord(session.toast.undoId);
            dismissToast();
          }}
        >
          {session.toast.text}{session.toast.undoId ? " · 되돌리기" : ""}
        </button>
      )}
    </div>
  );
}
