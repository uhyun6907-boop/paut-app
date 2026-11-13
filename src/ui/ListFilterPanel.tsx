/**
 * 리스트 탭 열 필터 화면: 머리글을 누르면 표 위에 작은 필터 칸이 열린다(검색 + 값 체크 + 이 열 해제).
 * 화면 위에 띄우지 않고 문서 흐름 안에 둬서 두 손가락으로 확대한 상태에서도 그대로 보이고 눌린다.
 * 앱 안 보기만 바꾼다 — 엑셀 자동 필터·행 숨김은 건드리지 않는다.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import type { LiveRow } from "../excel/liveList";
import { KEY_COLS } from "./listLayout";
import { activeColumns, clearColumn, columnValues, countText, filterableColumns, isActive, setText, toggleValue, type ListFilters } from "./listFilter";

/** 한 번에 그리는 값 수(값이 아주 많은 열은 검색으로 좁힌다). */
const MAX_VALUES = 300;

export function FilterBar(props: {
  header: { col: string; title: string }[];
  filters: ListFilters;
  shown: number;
  total: number;
  onOpen: (col: string) => void;
  onClearAll: () => void;
}) {
  const { header, filters, shown, total, onOpen, onClearAll } = props;
  const active = activeColumns(filters);
  const all = filterableColumns(header);
  // 주요 열(날짜·라인·포인트·판정·비고)과 지금 걸린 열은 바로, 나머지 열은 '다른 열' 에서 고른다.
  const quick = all.filter((h) => (KEY_COLS as readonly string[]).includes(h.col) || active.includes(h.col));
  const others = all.filter((h) => !quick.includes(h));
  return (
    <div className="filter-bar">
      <div className="filter-chips">
        <span className="k">필터</span>
        {quick.map((h) => (
          <button key={h.col} type="button" className={`chip${active.includes(h.col) ? " on" : ""}`} onClick={() => onOpen(h.col)}>
            {h.title}
            {active.includes(h.col) ? " ●" : " ▾"}
          </button>
        ))}
        {others.length > 0 && (
          <select
            className="chip chip-select"
            aria-label="다른 열 필터"
            value=""
            onChange={(e) => {
              if (e.target.value) onOpen(e.target.value);
            }}
          >
            <option value="">다른 열 ▾</option>
            {others.map((h) => (
              <option key={h.col} value={h.col}>{h.title}</option>
            ))}
          </select>
        )}
      </div>
      {active.length > 0 ? (
        <div className="filter-count" role="status">
          <span>필터 {active.length}개 · {countText(shown, total)}</span>
          <button type="button" className="chip clear" onClick={onClearAll}>필터 모두 해제</button>
        </div>
      ) : (
        <div className="filter-hint">열 이름을 누르면 그 열로 거를 수 있습니다 · {total.toLocaleString("ko-KR")}행</div>
      )}
    </div>
  );
}

export function FilterPanel(props: {
  col: string;
  title: string;
  rows: LiveRow[];
  filters: ListFilters;
  onChange: (next: ListFilters) => void;
  onClose: () => void;
}) {
  const { col, title, rows, filters, onChange, onClose } = props;
  const box = useRef<HTMLDivElement>(null);
  const [find, setFind] = useState("");
  const f = filters[col];
  const values = useMemo(() => columnValues(rows, col, filters), [rows, col, filters]);
  const all = values.map((v) => v.value);
  const q = find.trim().toLowerCase();
  const listed = q ? values.filter((v) => v.value.toLowerCase().includes(q)) : values;
  const checked = (v: string) => !f?.values || f.values.includes(v);

  useEffect(() => {
    setFind(f?.text ?? "");
    box.current?.scrollIntoView?.({ block: "nearest" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [col]);

  const set = (next: ListFilters[string]) => onChange(isActive(next) ? { ...filters, [col]: next } : clearColumn(filters, col));

  return (
    <div className="card filter-panel" ref={box} role="dialog" aria-label={`${title} 필터`}>
      <div className="filter-head">
        <strong>「{title}」 필터</strong>
        <button type="button" className="chip" onClick={onClose}>닫기</button>
      </div>
      <input
        className="field-input"
        type="search"
        placeholder="검색(이 글자가 들어간 칸만)"
        value={find}
        aria-label={`${title} 검색`}
        onChange={(e) => {
          setFind(e.target.value);
          set(setText(f, e.target.value));
        }}
      />
      <div className="filter-actions">
        <button type="button" className="chip" onClick={() => set({ values: null, text: f?.text ?? "" })}>모두 선택</button>
        <button type="button" className="chip" onClick={() => set({ values: [], text: f?.text ?? "" })}>모두 해제</button>
        <button type="button" className="chip clear" onClick={() => { setFind(""); onChange(clearColumn(filters, col)); }}>이 열 필터 지우기</button>
      </div>
      <div className="filter-values">
        {listed.slice(0, MAX_VALUES).map((v) => (
          <label key={v.value} className="filter-value">
            <input type="checkbox" checked={checked(v.value)} onChange={() => set(toggleValue(f, v.value, all))} />
            <span className="v">{v.value}</span>
            <span className="n">{v.count}</span>
          </label>
        ))}
        {listed.length > MAX_VALUES && <div className="filter-hint">값이 많아 {MAX_VALUES}개만 보입니다 — 검색으로 좁히세요.</div>}
        {listed.length === 0 && <div className="filter-hint">맞는 값이 없습니다.</div>}
      </div>
    </div>
  );
}
