/**
 * 칠하기(드래그 합격) 제스처. DOM 없이 테스트할 수 있게 순수 로직만 둔다.
 *
 * - 손가락이 카드 위에서 "거의 가로"로 MOVE_SLOP 보다 많이 움직이면 칠하기를 시작한다.
 * - 먼저 세로로 움직이면 스크롤로 보고 이 터치 동안은 칠하지 않는다.
 * - MOVE_SLOP 은 카드 길게 누르기(편집) 취소 거리와 같은 값이라, 칠하기가 시작되는 순간
 *   길게 누르기 타이머도 반드시 함께 꺼진다.
 */

/** 이 거리(px)를 넘게 움직이면 탭·길게 누르기가 아니다. PointCard 와 공유한다. */
export const MOVE_SLOP = 12;
/** 카드를 이만큼(ms) 누르고 있으면 길게 누르기(기록 카드 → 날짜 고치기). */
export const LONG_PRESS_MS = 500;

/**
 * 카드 하나의 길게 누르기 판별. DOM 없이 테스트할 수 있게 타이머만 쓴다.
 * - LONG_PRESS_MS 동안 MOVE_SLOP 안에서 누르고 있으면 onLong 을 한 번 부른다.
 * - 그 전에 손가락이 MOVE_SLOP 보다 움직이면(드래그 합격·스크롤) 취소된다.
 * - 길게 누르기가 성립한 뒤 오는 click 은 탭으로 치지 않는다(click() 이 false).
 */
export class LongPress {
  private timer: ReturnType<typeof setTimeout> | null = null;
  private x = 0;
  private y = 0;
  private fired = false;

  constructor(private readonly onLong: () => void, private readonly ms = LONG_PRESS_MS) {}

  down(x: number, y: number): void {
    this.stop();
    this.fired = false;
    this.x = x;
    this.y = y;
    this.timer = setTimeout(() => {
      this.timer = null;
      this.fired = true;
      this.onLong();
    }, this.ms);
  }

  move(x: number, y: number): void {
    if (Math.abs(x - this.x) > MOVE_SLOP || Math.abs(y - this.y) > MOVE_SLOP) this.stop();
  }

  up(): void {
    this.stop();
  }

  cancel(): void {
    this.stop();
  }

  /** click 이 왔을 때 탭으로 처리할지. 길게 누르기 직후의 click 은 한 번 버린다. */
  click(): boolean {
    if (this.fired) {
      this.fired = false;
      return false;
    }
    return true;
  }

  private stop(): void {
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
  }
}
/** 가로 이동이 세로 이동의 몇 배 이상이어야 칠하기로 보나. */
export const HORIZONTAL_RATIO = 1.5;

export type MoveKind = "pending" | "paint" | "scroll";

export function classifyMove(dx: number, dy: number): MoveKind {
  const ax = Math.abs(dx);
  const ay = Math.abs(dy);
  if (ax > MOVE_SLOP && ax >= ay * HORIZONTAL_RATIO) return "paint";
  if (ay > MOVE_SLOP || ax > MOVE_SLOP) return "scroll";
  return "pending";
}

export interface PaintCardInfo {
  kind: "record" | "draft";
  /** 기록 카드의 결과. 합격 / 불합격 / 없음 */
  result?: string;
  /** 회색 재검사 카드 */
  reinspect?: boolean;
}

/**
 * 지나간 카드를 합격으로 바꿀지.
 * - 미입력(회색) → 합격
 * - 이미 합격 → 그대로
 * - 없음 → 그대로 (용접이 없는 번호)
 * - 불합격, 회색 재검사 → 드래그를 그 카드에서 시작했을 때만 합격. 지나가기만 하면 건드리지 않는다.
 */
export function paintAction(card: PaintCardInfo, fromStart: boolean): "pass" | "skip" {
  if (card.kind === "draft") return card.reinspect && !fromStart ? "skip" : "pass";
  if (card.result === "없음" || card.result === "합격") return "skip";
  if (card.result === "불합격") return fromStart ? "pass" : "skip";
  return "skip";
}

export interface PaintHit {
  key: string;
  fromStart: boolean;
}

export interface PaintRelease {
  /** true 면 이어서 오는 click(탭)을 무시해야 한다. */
  swallowClick: boolean;
  /** 손을 뗀 순간 칠해져 있는 카드. 이것만 실제로 저장한다. */
  commit: PaintHit[];
}

/**
 * 한 번의 터치(누름→뗌) 동안의 상태.
 *
 * 지나간 카드는 "길(path)"로 쌓인다. 이미 길에 있는 카드로 손가락이 되돌아오면
 * 그 뒤의 카드는 길에서 빠진다(되돌리기). 시작 카드는 손가락이 처음 누른 자리
 * (MOVE_SLOP 안)로 돌아오면 빠진다. 저장은 손을 뗄 때 남은 길만 한다.
 */
export class PaintGesture {
  phase: "idle" | MoveKind = "idle";
  private x = 0;
  private y = 0;
  private startKey: string | null = null;
  private trail: string[] = [];
  private didPaint = false;

  get painting(): boolean {
    return this.phase === "paint";
  }

  /** 지금 미리보기로 칠해져 있는 카드(지나간 순서). */
  get path(): PaintHit[] {
    return this.trail.map((key) => ({ key, fromStart: key === this.startKey }));
  }

  down(x: number, y: number, key: string | null): void {
    this.reset();
    this.phase = "pending";
    this.x = x;
    this.y = y;
    this.startKey = key;
  }

  /** 손가락 아래 카드(key)를 받는다. 미리보기(path)가 바뀌었으면 true. */
  move(x: number, y: number, keyUnder: string | null): boolean {
    if (this.phase === "pending") {
      const kind = classifyMove(x - this.x, y - this.y);
      if (kind === "pending") return false;
      this.phase = kind;
      if (kind !== "paint") return false;
      this.didPaint = true;
      if (this.startKey) this.trail.push(this.startKey);
      if (keyUnder && !this.trail.includes(keyUnder)) this.trail.push(keyUnder);
      return this.trail.length > 0;
    }
    if (this.phase !== "paint") return false;
    const home = Math.abs(x - this.x) <= MOVE_SLOP && Math.abs(y - this.y) <= MOVE_SLOP;
    if (this.startKey && keyUnder === this.startKey && home) {
      if (!this.trail.length) return false;
      this.trail = [];
      return true;
    }
    if (!keyUnder) return false;
    const at = this.trail.indexOf(keyUnder);
    if (at === -1) {
      this.trail.push(keyUnder);
      return true;
    }
    if (at === this.trail.length - 1) return false;
    this.trail = this.trail.slice(0, at + 1);
    return true;
  }

  /** 손을 뗌. 남은 길을 돌려주고 상태를 비운다. */
  up(): PaintRelease {
    const release = { swallowClick: this.didPaint, commit: this.didPaint ? this.path : [] };
    this.reset();
    return release;
  }

  /** 스크롤로 넘어감(pointercancel) 또는 길게 누르기가 먼저 성립함. 아무것도 저장하지 않는다. */
  cancel(): void {
    this.reset();
  }

  private reset(): void {
    this.phase = "idle";
    this.startKey = null;
    this.trail = [];
    this.didPaint = false;
  }
}
