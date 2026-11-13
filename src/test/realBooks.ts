/** 테스트 전용: 실제 OneDrive 통합 문서의 복사본(src/excel/fixtures)과 가짜 드라이브 준비. */
import { AJEON_STATEMENT_FILE } from "../excel/xls/ajeonStatement";
import { readFileSync } from "node:fs";
import { defaultMappings, defaultWeldLengths } from "../domain/mapping";
import type { AppConfig, DayMeta, InspectionRecord } from "../domain/types";
import { LIST_GM_FILE } from "../domain/types";
import { parseListWorkbook } from "../excel/listBook";
import { unzipXlsx } from "../excel/ooxml";
import { FakeDrive } from "./fakeGraph";
import { userCreatesMonthSheet } from "./userSheet";

export const LIST_ID = "LIST-ITEM";
export const ILJI_ID = "ILJI-ITEM";
export const STMT_ID = "STMT-ITEM";
/** 군포·마전 리스트: 실제 OneDrive 항목 ID·이름(드라이브 루트) 그대로. */
export const LIST_GM_ID = LIST_GM_FILE.id;
export const LIST_GM_NAME = LIST_GM_FILE.name;

const cache = new Map<string, Uint8Array>();

export const AJEON_ID = AJEON_STATEMENT_FILE.id;

export function fixture(name: "real-list.xlsx" | "real-worklog.xlsx" | "real-statement.xlsx" | "real-list-gunpo-majeon.xlsx" | "ajeon-real-sheets-2611.xls" | "ajeon-lo-styled.xls"): Uint8Array {
  if (!cache.has(name)) cache.set(name, new Uint8Array(readFileSync(new URL(`../excel/fixtures/${name}`, import.meta.url))));
  return cache.get(name)!;
}

/**
 * 실제 세 파일을 담은 가짜 OneDrive.
 * novSheet(기본 true): 사용자가 거래명세서에 '드림합정26.11' 빈 시트를 이미 만들어 둔 상태.
 * 앱은 새 달 시트를 만들지 않으므로, 시트가 없는 상황은 { novSheet: false } 로 시험한다.
 */
export function realDrive(opts: { novSheet?: boolean } = {}): FakeDrive {
  const drive = new FakeDrive();
  drive.put(LIST_ID, "평택에너지(택지,브레인,합정주공) NDT 리스트.xlsx", fixture("real-list.xlsx"));
  drive.put(ILJI_ID, "PAUT 근무일지.xlsx", fixture("real-worklog.xlsx"));
  drive.put(STMT_ID, "02_PAUT 거래명세서 (지성, 드림).xlsx", fixture("real-statement.xlsx"));
  drive.put(LIST_GM_ID, LIST_GM_NAME, fixture("real-list-gunpo-majeon.xlsx"));
  // 아전 거래명세서(.xls): 실제 파일의 시트 구성(아전 군포 2611 포함)을 담은 로컬 사본. 실제 OneDrive 파일은 아니다.
  drive.putRaw(AJEON_ID, AJEON_STATEMENT_FILE.name, fixture("ajeon-real-sheets-2611.xls"));
  if (opts.novSheet !== false) userCreatesMonthSheet(drive.files(STMT_ID), "드림합정", 2026, 11);
  return drive;
}

let parsed: InspectionRecord[] | null = null;
let parsedGm: InspectionRecord[] | null = null;

/** 군포·마전 리스트를 읽은 기록. 매번 새 사본. */
export function realGmRecords(): InspectionRecord[] {
  if (!parsedGm) parsedGm = parseListWorkbook(unzipXlsx(fixture("real-list-gunpo-majeon.xlsx")), "listGm").flatMap((s) => s.records);
  return structuredClone(parsedGm);
}

/** 실제 리스트를 읽은 기록(11/11 앱 행 7개 포함). 매번 새 사본. */
export function realListRecords(): InspectionRecord[] {
  if (!parsed) parsed = parseListWorkbook(unzipXlsx(fixture("real-list.xlsx"))).flatMap((s) => s.records);
  return structuredClone(parsed);
}

export function graphConfig(): AppConfig {
  return {
    clientId: "test-client",
    syncMode: "graph",
    files: {
      list: { id: LIST_ID, name: "평택에너지(택지,브레인,합정주공) NDT 리스트.xlsx" },
      ilji: { id: ILJI_ID, name: "PAUT 근무일지.xlsx" },
      statement: { id: STMT_ID, name: "02_PAUT 거래명세서 (지성, 드림).xlsx" },
    },
    mappings: defaultMappings(),
    weldLengths: defaultWeldLengths(),
    inspectors: ["조우현", "심명보", "김정환"],
  };
}

export function meta(partial: Partial<DayMeta> = {}): DayMeta {
  return {
    workDate: "2026-11-12",
    dayType: "평일",
    insp1: "조우현",
    insp2: "심명보",
    timeStart: "09:00",
    timeEnd: "15:30",
    lastSite: "드림 합정주공",
    activeLine: "HJ-M-005-350A",
    prefix: "S",
    ...partial,
  };
}

let seq = 0;

export function rec(point: string, result: string, partial: Partial<InspectionRecord> = {}): InspectionRecord {
  seq += 1;
  return {
    id: `00000000-test-${String(seq).padStart(4, "0")}-${point}`,
    site: "드림 합정주공",
    sheetName: " 드림 합정주공",
    date: "2026-11-12",
    line: "HJ-M-005-350A",
    point,
    result,
    diameter: "350A",
    material: "SPPS380",
    sch: "",
    reportNo: "",
    reportDate: "",
    defectNo: "",
    drawingNo: "",
    weldNo: "",
    rtSheets: "",
    flawType: "",
    flawLen: "",
    note: "",
    issued: "",
    src: "user",
    sync: "pending",
    ts: seq,
    ...partial,
  };
}
