import { useState } from "react";
import { useDrill } from "./drill";
import { ILJI_COLUMNS, LIST_GM_FILE, type AppConfig, type SiteMapping } from "../domain/types";
import { listXlsxFiles, searchDrive, type DriveItem } from "../sync/graph";
import { accessToken, redirectUri, resolveClientId } from "../sync/msal";
import { resetSamples, saveConfig, signIn, signOut, useOneDriveFile, useSession } from "../store/session";
import type { BookKind } from "../excel/engine";
import { settingsAuthModel } from "../store/syncStatus";

/** 빌드 번호(커밋·시각). 설정 맨 아래에 보여 이 판이 깔렸는지 확인한다. */
export const BUILD_ID = import.meta.env.VITE_BUILD_ID || "dev";

const FILE_LABEL: Record<BookKind, string> = {
  list: "비파괴검사 리스트 (평택: 택지·브레인·합정주공·화양)",
  listGm: "비파괴검사 리스트 (지성 군포·아전 군포·지성 마전)",
  ilji: "작업일지",
  statement: "거래명세서",
};

const FILE_ORDER: readonly BookKind[] = ["list", "listGm", "ilji", "statement"];

/** 설정에 보일 파일 이름. 군포·마전 리스트는 따로 고르지 않았으면 드라이브 루트의 기본 파일. */
function fileName(config: AppConfig, kind: BookKind): string {
  if (kind === "listGm") {
    if (config.files.listGm === undefined) return `${LIST_GM_FILE.name} (드라이브 루트, 자동)`;
    return config.files.listGm?.name || "고르지 않음";
  }
  return config.files[kind]?.name || "데모 샘플 사용 중";
}

export function SettingsScreen() {
  const session = useSession();
  const [draft, setDraft] = useState<AppConfig>(session.config);
  const [picker, setPicker] = useState<BookKind | null>(null);
  useDrill(picker != null, () => setPicker(null));
  const [items, setItems] = useState<DriveItem[]>([]);
  const [query, setQuery] = useState("");
  const [pickError, setPickError] = useState("");
  const dirty = JSON.stringify(draft) !== JSON.stringify(session.config);
  const auth = settingsAuthModel({ account: session.account, auth: session.auth, loginNeeded: session.loginNeeded });

  function updateMap(site: string, partial: Partial<SiteMapping>) {
    setDraft({
      ...draft,
      mappings: draft.mappings.map((m) => (m.site === site ? { ...m, ...partial } : m)),
    });
  }

  async function openPicker(kind: BookKind) {
    setPicker(kind);
    setPickError("");
    setQuery("");
    setItems([]);
    try {
      const token = await accessToken(resolveClientId(draft.clientId));
      setItems(await listXlsxFiles(token));
    } catch (error) {
      setPickError(error instanceof Error ? error.message : "엑셀 목록을 열지 못했습니다. 먼저 Microsoft 로그인을 하세요.");
    }
  }

  async function choose(item: DriveItem) {
    if (!picker) return;
    if (item.folder) return;
    const next = { ...draft, syncMode: "graph" as const, files: { ...draft.files, [picker]: { id: item.id, name: item.name } } };
    setDraft(next);
    setPicker(null);
    await saveConfig(next);
    if (next.syncMode === "graph") await useOneDriveFile(picker);
  }

  return (
    <div>
      <div className="card">
        <div className="title">Microsoft 연결</div>
        <div className="row">
          <button className={`btn ${draft.syncMode === "demo" ? "primary" : ""}`} type="button" onClick={() => setDraft({ ...draft, syncMode: "demo" })}>데모</button>
          <button className={`btn ${draft.syncMode === "graph" ? "primary" : ""}`} type="button" onClick={() => setDraft({ ...draft, syncMode: "graph" })}>OneDrive</button>
        </div>
        <div className="k">Azure 클라이언트 ID</div>
        <input className="field-input" value={draft.clientId} placeholder="비우면 빌드 환경 변수 사용" onChange={(e) => setDraft({ ...draft, clientId: e.target.value.trim() })} style={{ marginTop: 4 }} />
        <p className="nav-hint">개인 Microsoft 계정. 권한 User.Read, Files.ReadWrite, offline_access. 로그인 후 돌아오는 주소는 {redirectUri()} 입니다. 클라이언트 ID를 비워 두면 기본값이 쓰입니다.</p>
        {auth.email ? (
          <>
            <div className="v">{auth.email}</div>
            <div className="k">{auth.status}</div>
            {auth.showRelogin && (
              <button className="btn primary block" type="button" onClick={() => void signIn()}>다시 로그인</button>
            )}
            <button className="btn block" type="button" onClick={() => void signIn({ switchAccount: true })}>다른 계정</button>
            <button className="btn block" type="button" onClick={() => void signOut()}>로그아웃</button>
          </>
        ) : (
          <>
            <button className="btn primary block" type="button" onClick={() => void signIn()}>Microsoft 로그인</button>
            <button className="btn block" type="button" onClick={() => void signIn({ switchAccount: true })}>다른 계정</button>
          </>
        )}
        <p className="nav-hint">한 번 로그인하면 이 기기에 기억됩니다. Microsoft 정책상 앱(SPA)의 갱신 토큰은 24시간마다 끝나므로, 하루 넘게 쓰지 않은 뒤에는 화면 위에 '로그인 필요'가 뜰 수 있습니다. 누르면 보통 비밀번호 없이 바로 돌아오고, 대기 기록은 그때 자동으로 보냅니다.</p>
      </div>

      <div className="card">
        <div className="title">OneDrive 통합 문서</div>
        {FILE_ORDER.map((kind) => (
          <div className="file-row" key={kind}>
            <div>
              <strong>{FILE_LABEL[kind]}</strong>
              <div className="sub" style={{ color: "var(--muted)", fontSize: 12 }}>{fileName(draft, kind)}</div>
            </div>
            <button className="btn small" type="button" onClick={() => void openPicker(kind)}>{draft.files[kind] || kind === "listGm" ? "바꾸기" : "엑셀 고르기"}</button>
          </div>
        ))}
        <p className="nav-hint">검사리스트는 두 파일입니다. 현장(시트)마다 자기 리스트 파일에만 씁니다: 지성 군포 PA·아전 군포 PA·지성 마전 PA 는 군포·마전 리스트, 나머지는 평택 리스트. 네 파일 모두 바뀐 칸만 씁니다(파일 전체를 덮어쓰지 않음). PC 엑셀에서 같은 파일을 편집 중이면 잠김으로 실패할 수 있습니다. 그때는 PC에서 저장하고 닫은 뒤 화면 위의 다시 시도를 누르세요.</p>
      </div>

      <div className="section-title">현장 → 작업일지 · 거래명세서</div>
      <p className="nav-hint">지성 브레인, 택지 드림, 드림 화양, 브레인대지 인입은 앱에서 지웠습니다(엑셀 시트·행은 그대로, 보내지 않음). 아전 군포 PA 의 명세서는 아직 쓰지 않습니다(건너뜀). 시트가 없는 접두어는 비워 두었고, 확인을 켜기 전에는 명세서를 만들지 않습니다.</p>
      {draft.mappings.map((m) => (
        <div className="card" key={m.site}>
          <div className="title" style={{ display: "flex", justifyContent: "space-between" }}>
            <span>{m.site}</span>
            {(!m.prefixConfirmed || m.unitPrice == null) && <span className="warn">확인 필요</span>}
          </div>
          <div className="k">작업일지 열</div>
          <select className="site-select" value={m.iljiColumn} onChange={(e) => updateMap(m.site, { iljiColumn: e.target.value })} style={{ margin: "4px 0 8px" }}>
            {ILJI_COLUMNS.map((col) => (
              <option key={col} value={col}>{col}</option>
            ))}
          </select>
          <div className="k">거래명세서 시트 접두어</div>
          <input className="field-input" value={m.statementPrefix} placeholder="예: 드림합정" onChange={(e) => updateMap(m.site, { statementPrefix: e.target.value })} style={{ marginTop: 4 }} />
          <label className="check">
            <input type="checkbox" checked={m.prefixConfirmed} onChange={(e) => updateMap(m.site, { prefixConfirmed: e.target.checked })} />
            이 접두어로 시트를 찾는다(시트는 직접 만든다)
          </label>
          <div className="k">단가 (원/m)</div>
          <input
            className="field-input"
            inputMode="numeric"
            value={m.unitPrice ?? ""}
            placeholder="비우면 금액을 쓰지 않음"
            onChange={(e) => {
              const raw = e.target.value.replace(/,/g, "").trim();
              updateMap(m.site, { unitPrice: raw === "" ? null : Number(raw), priceConfirmed: raw !== "" });
            }}
            style={{ marginTop: 4 }}
          />
          <div className="map-note">{m.priceNote}</div>
        </div>
      ))}

      <div className="card">
        <div className="title">관경별 용접길이 (m)</div>
        <div className="length-grid">
          {Object.keys(draft.weldLengths).sort((a, b) => Number(a) - Number(b)).map((key) => (
            <label key={key}>
              {key}A
              <input
                className="field-input"
                value={draft.weldLengths[key]}
                onChange={(e) => setDraft({ ...draft, weldLengths: { ...draft.weldLengths, [key]: Number(e.target.value) } })}
              />
            </label>
          ))}
        </div>
        <button
          className="btn block"
          type="button"
          onClick={() => {
            const key = window.prompt("관경 A 숫자 (예: 150)");
            if (!key || !/^\d+$/.test(key)) return;
            setDraft({ ...draft, weldLengths: { ...draft.weldLengths, [key]: 0 } });
          }}
        >
          관경 추가
        </button>
      </div>

      <button className="btn primary block" type="button" disabled={!dirty} onClick={() => void saveConfig(draft)}>설정 저장</button>
      <button className="btn danger block" type="button" onClick={() => void resetSamples()}>샘플 통합 문서로 되돌리기</button>

      <div className="card" style={{ marginTop: 16 }}>
        <div className="title">Azure 앱 등록 요약</div>
        <ol className="nav-hint" style={{ paddingLeft: 18, lineHeight: 1.5 }}>
          <li>portal.azure.com → Microsoft Entra ID → 앱 등록 → 새 등록.</li>
          <li>지원 계정: 개인 Microsoft 계정만.</li>
          <li>리디렉션 URI: SPA, {redirectUri()}</li>
          <li>API 권한: Microsoft Graph 위임 Files.ReadWrite, User.Read.</li>
          <li>클라이언트 암호는 만들지 않습니다. 애플리케이션 ID만 위에 넣습니다.</li>
        </ol>
        <p className="nav-hint">배포와 등록 절차는 README에 단계별로 있습니다.</p>
      </div>

      <p className="nav-hint build-id" style={{ textAlign: "center" }}>빌드 {BUILD_ID}</p>

      <div className={`modal-backdrop ${picker ? "show" : ""}`} onClick={() => setPicker(null)}>
        <div className="modal" onClick={(e) => e.stopPropagation()}>
          <h3>{picker ? FILE_LABEL[picker] : ""} 고르기</h3>
          <p className="nav-hint">OneDrive의 .xlsx를 누르면 이 칸에 기억됩니다.</p>
          <div className="row">
            <input className="field-input" value={query} placeholder="이름 검색" onChange={(e) => setQuery(e.target.value)} />
            <button
              className="btn"
              type="button"
              onClick={() => {
                void (async () => {
                  try {
                    const token = await accessToken(resolveClientId(draft.clientId));
                    const found = query.trim() ? await searchDrive(token, query.trim()) : await listXlsxFiles(token);
                    setItems(found.filter((item) => !item.folder && item.name.toLowerCase().endsWith(".xlsx")));
                    setPickError("");
                  } catch (error) {
                    setPickError(error instanceof Error ? error.message : "검색 실패");
                  }
                })();
              }}
            >
              검색
            </button>
          </div>
          {pickError && <div className="err">{pickError}</div>}
          {items.length === 0 && !pickError && <div className="empty-note">xlsx가 없습니다. 검색어를 바꿔 보세요.</div>}
          {items.map((item) => (
            <button key={item.id} className="choice" type="button" onClick={() => void choose(item)}>
              {item.name}
            </button>
          ))}
          <button className="btn block" type="button" onClick={() => setPicker(null)}>닫기</button>
        </div>
      </div>
    </div>
  );
}
