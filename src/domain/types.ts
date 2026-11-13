export type ResultKind = "합격" | "불합격";

export type SyncStatus = "local" | "pending" | "syncing" | "synced" | "error";

/** 동기화 대상 파일. 각각 따로 성공·실패한다. */
export type SyncFile = "list" | "ilji" | "statement";

export const SYNC_FILES: readonly SyncFile[] = ["list", "ilji", "statement"];

export const SYNC_FILE_LABEL: Record<SyncFile, string> = {
  list: "검사리스트",
  ilji: "작업일지",
  statement: "거래명세서",
};

/**
 * 검사리스트 통합 문서. 현장(시트)마다 들어갈 리스트 파일이 정해져 있다.
 * - list: 평택에너지(택지,브레인,합정주공) NDT 리스트.xlsx
 * - listGm: PAUT NDT 리스트(군포, 마전).xlsx (드라이브 루트)
 * 기록의 steps.list 는 "자기 리스트 파일에 반영됨"을 뜻한다(어느 파일인지는 listBookOf).
 */
export type ListBookKind = "list" | "listGm";

export const LIST_BOOKS: readonly ListBookKind[] = ["list", "listGm"];

export const LIST_BOOK_LABEL: Record<ListBookKind, string> = {
  list: "검사리스트(평택)",
  listGm: "검사리스트(군포·마전)",
};

/** 군포·마전 리스트(OneDrive 드라이브 루트). 설정에서 다른 파일을 고르면 그 값이 우선. */
export const LIST_GM_FILE = {
  id: "73DF94A06E5230E0!sa2c63c1f3c464937afdfb19a02c526df",
  name: "PAUT NDT 리스트(군포, 마전).xlsx",
} as const;

/** 군포·마전 리스트 파일의 PA 시트. 파일을 아직 받지 않았어도 이 시트의 기록은 그 파일로 보낸다. */
export const LIST_GM_SHEETS: readonly string[] = ["지성 군포 PA", "아전 군포 PA", "지성 마전 PA"];

/** 기록 하나가 파일 하나에 반영된 상태. 없으면 아직 안 보냄. */
export interface FileStep {
  /** wait: 실패가 아니라 기다림(예: 거래명세서에 그 달 시트가 아직 없음). 다음 동기화 때 다시 쓴다. */
  status: "done" | "error" | "skip" | "wait";
  /** 실패 이유(한국어), 건너뛴 이유 또는 기다리는 이유. */
  message?: string;
  at?: number;
}

export type FileSteps = Partial<Record<SyncFile, FileStep>>;

/** 마지막 동기화 결과. IndexedDB 에 남아 앱을 다시 켜도 보인다. */
export interface SyncReport {
  at: number;
  offline?: boolean;
  files: Record<SyncFile, { status: "ok" | "error" | "skip" | "idle"; message?: string; count?: number }>;
  /** 검사리스트 파일별 결과(이번에 다룬 파일만). files.list 는 두 리스트를 합친 결과. */
  lists?: Partial<Record<ListBookKind, { status: "ok" | "error" | "skip" | "idle"; message?: string; count?: number }>>;
  /** 아전 거래명세서(.xls) 결과(이번에 다룬 경우만). 아전 기록의 단계는 files.statement 와 같은 'statement'. */
  ajeon?: { status: "ok" | "error" | "skip" | "idle"; message?: string; count?: number };
}

/** 그날 작업일지에 쓸 검사자·근무시간. 저장할 때 날짜별로 남긴다. */
export interface DayStaff {
  dayType: "평일" | "공휴일";
  insp1: string;
  insp2: string;
  timeStart: string;
  timeEnd: string;
}

export interface InspectionRecord {
  id: string;
  site: string;
  sheetName: string;
  date: string;
  line: string;
  point: string;
  result: ResultKind | string;
  /** 청구·개소에 쓰는 관경. 라인에서 마지막 `숫자A`를 우선합니다. */
  diameter: string;
  material: string;
  sch: string;
  reportNo: string;
  reportDate: string;
  defectNo: string;
  drawingNo: string;
  weldNo: string;
  rtSheets: string;
  flawType: string;
  flawLen: string;
  note: string;
  issued: string;
  src: "xlsx" | "user";
  /** 이 기록이 들어갈 검사리스트 파일. 없으면 시트 이름으로 정한다(listBookOf). */
  listBook?: ListBookKind;
  excelRow?: number;
  sync: SyncStatus;
  syncError?: string;
  /** 파일별 반영 상태. 리스트는 됐는데 작업일지가 실패하면 작업일지만 다시 보낸다. */
  steps?: FileSteps;
  /** 거래명세서를 건너뛴 이유 등 */
  syncNote?: string;
  deleted?: boolean;
  /**
   * 카드를 길게 눌러 날짜를 바꾼 뒤 아직 엑셀에 다 반영되지 않았을 때, 바꾸기 전 날짜들.
   * 동기화가 그 날짜(작업일지)·그 달(거래명세서)도 다시 계산해 옛 자리에 남은 개수를 지운다. 반영되면 지운다.
   */
  movedFrom?: string[];
  ts: number;
}

export interface SiteBundle {
  site: string;
  sheetName: string;
  records: InspectionRecord[];
}

export interface DayMeta {
  workDate: string;
  dayType: "평일" | "공휴일";
  insp1: string;
  /** 빈 문자열이면 혼자 검사. 작업일지 둘째 칸은 비운다. */
  insp2: string;
  /** 날짜별 검사자 2. 빈 문자열은 혼자. */
  insp2ByDate?: Record<string, string>;
  timeStart: string;
  timeEnd: string;
  lastSite: string;
  activeLine: string;
  prefix: string;
  /** 라인별 진행 방향. 키는 `현장||라인`. */
  directionByLine?: Record<string, "up" | "down">;
  /** 날짜별로 저장 당시의 검사자·근무시간. 다른 날 기록을 나중에 보낼 때 그날 값을 쓴다. */
  dayStaff?: Record<string, DayStaff>;
  /** 아직 저장하지 않은 회색 카드. 키는 `현장||라인||접두어`. */
  draftCards?: Record<string, DraftCard[]>;
}

/** 저장 전 카드. reinspectOf 가 있으면 그 불합격 행의 재검사(같은 포인트 번호). */
export interface DraftCard {
  id: string;
  point: string;
  reinspectOf?: string;
}

export interface SiteMapping {
  site: string;
  /** 작업일지 관경/개소 열 이름. 없는 현장은 '그 외'. */
  iljiColumn: string;
  /** 거래명세서 시트 이름 앞부분. 비어 있으면 시트를 만들지 않습니다. */
  statementPrefix: string;
  /** 사용자가 접두어를 확인했거나, 파일에서 시트를 찾았으면 true. */
  prefixConfirmed: boolean;
  /** 원/M. null 이면 금액을 쓰지 않습니다. */
  unitPrice: number | null;
  priceConfirmed: boolean;
  priceNote: string;
}

export interface AppConfig {
  clientId: string;
  syncMode: "demo" | "graph";
  files: {
    list: { id: string; name: string } | null;
    /** 군포·마전 리스트. 예전 설정에는 없다(undefined → 기본 파일 LIST_GM_FILE). */
    listGm?: { id: string; name: string } | null;
    ilji: { id: string; name: string } | null;
    statement: { id: string; name: string } | null;
  };
  mappings: SiteMapping[];
  /** 관경(A) → 용접길이(m) */
  weldLengths: Record<string, number>;
  inspectors: string[];
  /** 목록에 직접 추가한 검사자. */
  addedInspectors?: string[];
}

export interface BillLine {
  date: string;
  diameter: string;
  count: number;
  lengthM: number | null;
  qtyM: number | null;
  unitPrice: number;
  amount: number | null;
}

export interface Bill {
  lines: BillLine[];
  totalM: number;
  supply: number;
  vat: number;
  grand: number;
  missingLengths: string[];
}

export const ILJI_COLUMNS = [
  "지성 브레인연계",
  "지성 마전",
  "지성 군포",
  "드림 합정주공",
  "아전 군포",
  "그 외",
] as const;

export type IljiColumn = (typeof ILJI_COLUMNS)[number];

export const PASS = "합격";
export const FAIL = "불합격";
