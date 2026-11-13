/**
 * OLE 복합 문서(CFB, .xls 컨테이너) 읽기·스트림 바꾸기.
 *
 * - 읽기는 엄격하다. 체인 고리·범위 밖 섹터·두 스트림이 같은 섹터를 쓰는 경우·크기와 체인 길이 불일치는
 *   모두 오류로 막는다(깨진 파일을 고쳐 쓰다가 데이터를 망치지 않도록).
 * - 바꾸기는 지정한 스트림의 섹터만 다시 쓴다. 헤더·FAT 의 그 스트림 항목·디렉터리의 그 스트림 크기 말고는
 *   다른 바이트(다른 스트림, 요약 정보, 디렉터리 나머지)를 그대로 둔다. 커지면 파일 끝에 섹터를 붙인다.
 */

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
export const FREESECT = 0xffffffff;
export const ENDOFCHAIN = 0xfffffffe;
export const FATSECT = 0xfffffffd;
export const DIFSECT = 0xfffffffc;
const MAXREGSECT = 0xfffffffa;

export class CfbError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "CfbError";
  }
}

export interface CfbEntry {
  index: number;
  name: string;
  /** 1 storage, 2 stream, 5 root */
  type: number;
  start: number;
  size: number;
}

export interface Cfb {
  bytes: Uint8Array;
  major: number;
  sectorSize: number;
  miniCutoff: number;
  sectorCount: number;
  fatSectors: number[];
  difatSectors: number[];
  fat: number[];
  dirChain: number[];
  entries: CfbEntry[];
}

function u16(b: Uint8Array, o: number): number {
  return b[o] | (b[o + 1] << 8);
}
function u32(b: Uint8Array, o: number): number {
  return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16)) + b[o + 3] * 0x1000000;
}
function put32(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
  b[o + 2] = (v >>> 16) & 0xff;
  b[o + 3] = (v >>> 24) & 0xff;
}

function sectorOffset(cfb: { sectorSize: number }, id: number): number {
  return (id + 1) * cfb.sectorSize;
}

/** 엄격하게 읽는다. 조금이라도 이상하면 CfbError. */
export function readCfb(bytes: Uint8Array): Cfb {
  if (bytes.length < 512 || SIGNATURE.some((v, i) => bytes[i] !== v)) throw new CfbError("OLE(.xls) 파일이 아닙니다.");
  if (u16(bytes, 28) !== 0xfffe) throw new CfbError("OLE 바이트 순서 표시가 이상합니다.");
  const major = u16(bytes, 26);
  const shift = u16(bytes, 30);
  if (!((major === 3 && shift === 9) || (major === 4 && shift === 12))) throw new CfbError(`지원하지 않는 OLE 버전입니다(${major}/${shift}).`);
  if (u16(bytes, 32) !== 6) throw new CfbError("OLE 미니 섹터 크기가 이상합니다.");
  const sectorSize = 1 << shift;
  const miniCutoff = u32(bytes, 56);
  if (miniCutoff !== 4096) throw new CfbError("OLE 미니 스트림 기준이 이상합니다.");
  if ((bytes.length - sectorSize) % sectorSize !== 0) throw new CfbError("OLE 파일 길이가 섹터 크기와 맞지 않습니다.");
  const sectorCount = (bytes.length - sectorSize) / sectorSize;
  const nFat = u32(bytes, 44);
  const firstDir = u32(bytes, 48);
  const firstDifat = u32(bytes, 68);
  const nDifat = u32(bytes, 72);
  const perSector = sectorSize / 4;

  const fatSectors: number[] = [];
  for (let i = 0; i < Math.min(nFat, 109); i++) fatSectors.push(u32(bytes, 76 + i * 4));
  for (let i = Math.min(nFat, 109); i < 109; i++) {
    if (u32(bytes, 76 + i * 4) !== FREESECT) throw new CfbError("OLE 헤더 DIFAT 가 이상합니다.");
  }
  const difatSectors: number[] = [];
  let next = firstDifat;
  for (let n = 0; n < nDifat; n++) {
    if (next >= sectorCount) throw new CfbError("OLE DIFAT 체인이 깨졌습니다.");
    difatSectors.push(next);
    const off = sectorOffset({ sectorSize }, next);
    for (let i = 0; i < perSector - 1 && fatSectors.length < nFat; i++) fatSectors.push(u32(bytes, off + i * 4));
    next = u32(bytes, off + (perSector - 1) * 4);
  }
  if (fatSectors.length !== nFat) throw new CfbError("OLE FAT 섹터 수가 맞지 않습니다.");
  const fat: number[] = [];
  for (const s of fatSectors) {
    if (s >= sectorCount) throw new CfbError("OLE FAT 섹터가 파일 밖을 가리킵니다.");
    const off = sectorOffset({ sectorSize }, s);
    for (let i = 0; i < perSector; i++) fat.push(u32(bytes, off + i * 4));
  }
  if (fat.length < sectorCount) throw new CfbError("OLE FAT 이 파일보다 짧습니다.");
  for (const s of fatSectors) if (fat[s] !== FATSECT) throw new CfbError("OLE FAT 섹터 표시가 이상합니다.");
  for (const s of difatSectors) if (fat[s] !== DIFSECT) throw new CfbError("OLE DIFAT 섹터 표시가 이상합니다.");
  for (let i = sectorCount; i < fat.length; i++) if (fat[i] !== FREESECT) throw new CfbError("OLE FAT 이 파일 밖 섹터를 씁니다.");

  const owner = new Map<number, string>();
  for (const s of fatSectors) owner.set(s, "FAT");
  for (const s of difatSectors) {
    if (owner.has(s)) throw new CfbError("OLE 섹터가 겹칩니다(DIFAT).");
    owner.set(s, "DIFAT");
  }
  const walk = (start: number, label: string): number[] => {
    const out: number[] = [];
    let s = start;
    while (s !== ENDOFCHAIN) {
      if (s > MAXREGSECT || s >= sectorCount) throw new CfbError(`OLE ${label} 체인이 깨졌습니다.`);
      if (owner.has(s)) throw new CfbError(`OLE 섹터가 겹칩니다(${label} ↔ ${owner.get(s)}).`);
      owner.set(s, label);
      out.push(s);
      if (out.length > sectorCount) throw new CfbError(`OLE ${label} 체인이 끝나지 않습니다.`);
      s = fat[s];
    }
    return out;
  };

  const cfb: Cfb = { bytes, major, sectorSize, miniCutoff, sectorCount, fatSectors, difatSectors, fat, dirChain: [], entries: [] };
  cfb.dirChain = walk(firstDir, "디렉터리");
  const entries: CfbEntry[] = [];
  for (let k = 0; k < cfb.dirChain.length * (sectorSize / 128); k++) {
    const off = sectorOffset(cfb, cfb.dirChain[Math.floor(k / (sectorSize / 128))]) + (k % (sectorSize / 128)) * 128;
    const type = bytes[off + 66];
    if (type === 0) continue;
    const nameLen = u16(bytes, off + 64);
    if (nameLen < 2 || nameLen > 64 || nameLen % 2) throw new CfbError("OLE 디렉터리 이름이 이상합니다.");
    let name = "";
    for (let i = 0; i < nameLen - 2; i += 2) name += String.fromCharCode(u16(bytes, off + i));
    const sizeHigh = u32(bytes, off + 124);
    if (major === 4 && sizeHigh !== 0) throw new CfbError("OLE 스트림이 너무 큽니다.");
    entries.push({ index: k, name, type, start: u32(bytes, off + 116), size: u32(bytes, off + 120) });
  }
  cfb.entries = entries;
  const root = entries.find((e) => e.type === 5);
  if (!root || root.index !== 0) throw new CfbError("OLE 루트 항목이 없습니다.");
  // 큰 스트림 체인(그리고 미니 스트림 자체)을 모두 따라가 겹침·크기를 확인한다.
  if (root.size > 0) {
    const chain = walk(root.start, "미니 스트림");
    if (chain.length !== Math.ceil(root.size / sectorSize)) throw new CfbError("OLE 미니 스트림 크기가 맞지 않습니다.");
  }
  const nMiniFat = u32(bytes, 64);
  if (nMiniFat > 0) {
    const chain = walk(u32(bytes, 60), "미니 FAT");
    if (chain.length !== nMiniFat) throw new CfbError("OLE 미니 FAT 크기가 맞지 않습니다.");
  }
  for (const e of entries) {
    if (e.type !== 2 || e.size < miniCutoff) continue;
    const chain = walk(e.start, e.name);
    if (chain.length !== Math.ceil(e.size / sectorSize)) throw new CfbError(`OLE '${e.name}' 스트림 크기와 체인 길이가 맞지 않습니다.`);
  }
  return cfb;
}

export function chainOf(cfb: Cfb, start: number): number[] {
  const out: number[] = [];
  let s = start;
  while (s !== ENDOFCHAIN) {
    out.push(s);
    s = cfb.fat[s];
  }
  return out;
}

export function findEntry(cfb: Cfb, name: string): CfbEntry | undefined {
  return cfb.entries.find((e) => e.type === 2 && e.name.toLowerCase() === name.toLowerCase());
}

/** 큰 스트림(미니 기준 이상)의 바이트. */
export function readStream(cfb: Cfb, name: string): Uint8Array {
  const e = findEntry(cfb, name);
  if (!e) throw new CfbError(`OLE 에 '${name}' 스트림이 없습니다.`);
  if (e.size < cfb.miniCutoff) throw new CfbError(`'${name}' 스트림이 너무 작습니다(미니 스트림).`);
  const out = new Uint8Array(e.size);
  let pos = 0;
  for (const s of chainOf(cfb, e.start)) {
    const off = sectorOffset(cfb, s);
    const n = Math.min(cfb.sectorSize, e.size - pos);
    out.set(cfb.bytes.subarray(off, off + n), pos);
    pos += n;
  }
  return out;
}

/** 모든 스트림 항목의 내용(미니 스트림 포함). 검증용. */
export function readAllStreams(cfb: Cfb): Map<string, Uint8Array> {
  const out = new Map<string, Uint8Array>();
  const root = cfb.entries.find((e) => e.type === 5)!;
  let mini: Uint8Array | null = null;
  const miniFat: number[] = [];
  const nMiniFat = u32(cfb.bytes, 64);
  if (nMiniFat > 0) {
    for (const s of chainOf(cfb, u32(cfb.bytes, 60))) {
      const off = sectorOffset(cfb, s);
      for (let i = 0; i < cfb.sectorSize / 4; i++) miniFat.push(u32(cfb.bytes, off + i * 4));
    }
  }
  for (const e of cfb.entries) {
    if (e.type !== 2) continue;
    if (e.size >= cfb.miniCutoff) {
      out.set(e.name, readStream(cfb, e.name));
      continue;
    }
    if (!mini) {
      mini = new Uint8Array(root.size);
      let pos = 0;
      for (const s of chainOf(cfb, root.start)) {
        const off = sectorOffset(cfb, s);
        const n = Math.min(cfb.sectorSize, root.size - pos);
        mini.set(cfb.bytes.subarray(off, off + n), pos);
        pos += n;
      }
    }
    const data = new Uint8Array(e.size);
    let pos = 0;
    let s = e.start;
    let guard = 0;
    while (pos < e.size) {
      if (s >= miniFat.length || guard++ > miniFat.length) throw new CfbError(`OLE 미니 체인이 깨졌습니다(${e.name}).`);
      const n = Math.min(64, e.size - pos);
      data.set(mini.subarray(s * 64, s * 64 + n), pos);
      pos += n;
      s = miniFat[s];
    }
    out.set(e.name, data);
  }
  return out;
}

/** 디렉터리 항목 바이트(128). 검증용. */
export function entryBytes(cfb: Cfb, index: number): Uint8Array {
  const per = cfb.sectorSize / 128;
  const off = sectorOffset(cfb, cfb.dirChain[Math.floor(index / per)]) + (index % per) * 128;
  return cfb.bytes.subarray(off, off + 128);
}

/**
 * 큰 스트림 하나의 내용을 바꾼 새 파일. 그 스트림의 섹터만 다시 쓰고,
 * 모자라면 파일 끝에 섹터를 붙인다(필요하면 FAT 섹터도 붙인다). 남는 섹터는 비움(FREESECT)으로 돌린다.
 */
export function replaceStream(src: Uint8Array, name: string, data: Uint8Array): Uint8Array {
  const cfb = readCfb(src);
  const entry = findEntry(cfb, name);
  if (!entry) throw new CfbError(`OLE 에 '${name}' 스트림이 없습니다.`);
  if (entry.size < cfb.miniCutoff || data.length < cfb.miniCutoff) throw new CfbError("미니 스트림 크기의 스트림은 바꾸지 않습니다.");
  const ss = cfb.sectorSize;
  const per = ss / 4;
  const oldChain = chainOf(cfb, entry.start);
  const need = Math.ceil(data.length / ss);
  const fat = cfb.fat.slice();
  let fatSectors = cfb.fatSectors.slice();
  let sectorCount = cfb.sectorCount;
  const chain = oldChain.slice(0, need);
  const appended: number[] = [];
  const newFatSectors: number[] = [];
  while (chain.length < need) {
    // 파일 끝에 섹터를 붙인다. FAT 이 모자라면 FAT 섹터를 먼저 붙인다.
    if (sectorCount + 1 > fatSectors.length * per) {
      if (fatSectors.length + 1 > 109) throw new CfbError("파일이 커서 OLE FAT 을 늘릴 수 없습니다.");
      const id = sectorCount++;
      fatSectors = [...fatSectors, id];
      newFatSectors.push(id);
      while (fat.length < fatSectors.length * per) fat.push(FREESECT);
      fat[id] = FATSECT;
      continue;
    }
    const id = sectorCount++;
    while (fat.length <= id) fat.push(FREESECT);
    chain.push(id);
    appended.push(id);
  }
  // 체인 다시 잇기
  for (let i = 0; i < chain.length; i++) fat[chain[i]] = i + 1 < chain.length ? chain[i + 1] : ENDOFCHAIN;
  const freed = oldChain.slice(need);
  for (const s of freed) fat[s] = FREESECT;

  const out = new Uint8Array((sectorCount + 1) * ss);
  out.set(src.subarray(0, src.length));
  // 스트림 내용
  for (let i = 0; i < chain.length; i++) {
    const off = sectorOffset(cfb, chain[i]);
    out.fill(0, off, off + ss);
    out.set(data.subarray(i * ss, Math.min(data.length, (i + 1) * ss)), off);
  }
  for (const s of freed) {
    const off = sectorOffset(cfb, s);
    out.fill(0, off, off + ss);
  }
  // FAT 섹터 다시 쓰기(바뀐 항목만 바뀐다)
  for (let k = 0; k < fatSectors.length; k++) {
    const off = sectorOffset(cfb, fatSectors[k]);
    for (let i = 0; i < per; i++) put32(out, off + i * 4, fat[k * per + i] ?? FREESECT);
  }
  // 헤더: FAT 섹터 수와 DIFAT
  put32(out, 44, fatSectors.length);
  for (let i = 0; i < 109; i++) put32(out, 76 + i * 4, i < fatSectors.length ? fatSectors[i] : FREESECT);
  if (cfb.difatSectors.length) throw new CfbError("DIFAT 섹터가 있는 큰 OLE 파일은 지원하지 않습니다.");
  // 디렉터리 항목 크기·시작 섹터
  const per128 = ss / 128;
  const eoff = sectorOffset(cfb, cfb.dirChain[Math.floor(entry.index / per128)]) + (entry.index % per128) * 128;
  put32(out, eoff + 116, chain[0]);
  put32(out, eoff + 120, data.length);
  if (cfb.major === 4) put32(out, eoff + 124, 0);
  void appended;
  void newFatSectors;
  return out;
}
