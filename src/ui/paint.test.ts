import { afterEach, describe, expect, it, vi } from "vitest";
import { LONG_PRESS_MS, LongPress, MOVE_SLOP, PaintGesture, classifyMove, paintAction } from "./paint";

describe("칠하기 제스처 판별", () => {
  it("작게 움직이면 아직 모름(탭·길게 누르기 그대로)", () => {
    expect(classifyMove(5, 3)).toBe("pending");
    expect(classifyMove(MOVE_SLOP, 0)).toBe("pending");
  });

  it("거의 가로로 움직이면 칠하기", () => {
    expect(classifyMove(MOVE_SLOP + 1, 0)).toBe("paint");
    expect(classifyMove(-20, 10)).toBe("paint");
  });

  it("세로나 대각선이면 스크롤로 넘긴다", () => {
    expect(classifyMove(0, 20)).toBe("scroll");
    expect(classifyMove(15, 14)).toBe("scroll");
    expect(classifyMove(3, -30)).toBe("scroll");
  });

  it("칠하기 시작 거리는 길게 누르기 취소 거리보다 짧지 않다 (칠하는 중에 편집 버튼이 뜨지 않음)", () => {
    // PointCard 는 MOVE_SLOP 을 넘으면 길게 누르기 타이머를 끈다. 칠하기도 MOVE_SLOP 을 넘어야 시작.
    for (let dx = 0; dx <= MOVE_SLOP; dx++) expect(classifyMove(dx, 0)).not.toBe("paint");
  });
});

describe("어떤 카드를 칠하나", () => {
  it("미입력은 합격, 이미 합격·없음은 그대로", () => {
    expect(paintAction({ kind: "draft" }, false)).toBe("pass");
    expect(paintAction({ kind: "record", result: "합격" }, false)).toBe("skip");
    expect(paintAction({ kind: "record", result: "없음" }, true)).toBe("skip");
  });

  it("불합격·회색 재검사는 시작 카드일 때만 합격", () => {
    expect(paintAction({ kind: "record", result: "불합격" }, false)).toBe("skip");
    expect(paintAction({ kind: "record", result: "불합격" }, true)).toBe("pass");
    expect(paintAction({ kind: "draft", reinspect: true }, false)).toBe("skip");
    expect(paintAction({ kind: "draft", reinspect: true }, true)).toBe("pass");
  });
});

describe("PaintGesture", () => {
  const keys = (g: PaintGesture) => g.path.map((hit) => hit.key);

  it("가로로 끌면 시작 카드부터 지나간 카드가 길에 쌓이고, 뗀 뒤 탭을 막는다", () => {
    const g = new PaintGesture();
    g.down(0, 0, "a");
    expect(g.move(4, 0, "a")).toBe(false);
    expect(g.painting).toBe(false);
    expect(g.move(20, 2, "b")).toBe(true);
    expect(g.path).toEqual([
      { key: "a", fromStart: true },
      { key: "b", fromStart: false },
    ]);
    expect(g.move(30, 3, "b")).toBe(false);
    // 한 번 칠하기가 시작되면 아래 줄로 내려가도 계속 칠한다.
    expect(g.move(30, 60, "c")).toBe(true);
    expect(g.move(30, 60, null)).toBe(false);
    expect(keys(g)).toEqual(["a", "b", "c"]);
    const release = g.up();
    expect(release.swallowClick).toBe(true);
    expect(release.commit.map((hit) => hit.key)).toEqual(["a", "b", "c"]);
    expect(g.phase).toBe("idle");
  });

  it("되돌아가면 그 뒤로 칠한 카드는 길에서 빠지고, 다시 가면 다시 칠한다", () => {
    const g = new PaintGesture();
    g.down(0, 0, "s23");
    ["s23", "s24", "s25", "s26", "s27", "s28"].forEach((key, i) => g.move(14 + i * 44, 0, key));
    expect(keys(g)).toEqual(["s23", "s24", "s25", "s26", "s27", "s28"]);
    expect(g.move(14 + 2 * 44, 0, "s25")).toBe(true);
    expect(keys(g)).toEqual(["s23", "s24", "s25"]);
    expect(g.move(14 + 3 * 44, 0, "s26")).toBe(true);
    expect(keys(g)).toEqual(["s23", "s24", "s25", "s26"]);
    expect(g.up().commit.map((hit) => hit.key)).toEqual(["s23", "s24", "s25", "s26"]);
  });

  it("시작 카드까지 돌아와도 시작 카드는 남고, 처음 누른 자리(12px 안)로 돌아오면 시작 카드도 빠진다", () => {
    const g = new PaintGesture();
    g.down(0, 0, "a");
    g.move(20, 0, "a");
    g.move(60, 0, "b");
    g.move(-15, 0, "a");
    expect(keys(g)).toEqual(["a"]);
    expect(g.move(5, 0, "a")).toBe(true);
    expect(keys(g)).toEqual([]);
    // 다시 밀면 다시 칠한다.
    expect(g.move(20, 0, "a")).toBe(true);
    expect(keys(g)).toEqual(["a"]);
    g.move(3, 0, "a");
    const release = g.up();
    expect(release.commit).toEqual([]);
    expect(release.swallowClick).toBe(true);
  });

  it("세로로 먼저 움직이면 이번 터치는 칠하지 않는다 (스크롤)", () => {
    const g = new PaintGesture();
    g.down(0, 0, "a");
    expect(g.move(2, 25, "a")).toBe(false);
    expect(g.phase).toBe("scroll");
    expect(g.move(80, 25, "b")).toBe(false);
    expect(g.up()).toEqual({ swallowClick: false, commit: [] });
  });

  it("그냥 탭은 칠하지 않고 탭도 막지 않는다", () => {
    const g = new PaintGesture();
    g.down(0, 0, "a");
    expect(g.move(3, 2, "a")).toBe(false);
    expect(g.up()).toEqual({ swallowClick: false, commit: [] });
  });

  it("길게 누르기(편집)가 먼저 성립하면 이후 움직임은 칠하지 않는다", () => {
    const g = new PaintGesture();
    g.down(0, 0, "a");
    g.cancel(); // onLong 에서 호출
    expect(g.move(40, 0, "b")).toBe(false);
    expect(g.up()).toEqual({ swallowClick: false, commit: [] });
  });

  it("빈틈에서 시작해도 지나간 카드만 칠한다", () => {
    const g = new PaintGesture();
    g.down(0, 0, null);
    g.move(20, 0, "b");
    expect(g.path).toEqual([{ key: "b", fromStart: false }]);
  });

  it("스크롤로 넘어가 취소(pointercancel)되면 아무것도 저장하지 않는다", () => {
    const g = new PaintGesture();
    g.down(0, 0, "a");
    g.move(20, 0, "a");
    g.cancel();
    expect(g.move(60, 0, "b")).toBe(false);
    expect(g.up()).toEqual({ swallowClick: false, commit: [] });
  });
});

describe("카드 길게 누르기(날짜 고치기) vs 탭 vs 드래그", () => {
  afterEach(() => vi.useRealTimers());

  /** 한 카드 위의 터치: 카드(LongPress)와 격자(PaintGesture)가 같은 포인터를 함께 받는다. */
  function touch() {
    vi.useFakeTimers();
    const longs: number[] = [];
    const taps: number[] = [];
    const press = new LongPress(() => longs.push(Date.now()));
    const paint = new PaintGesture();
    return {
      longs,
      taps,
      paint,
      down(x = 100, y = 100) {
        press.down(x, y);
        paint.down(x, y, "S01");
      },
      move(x: number, y: number, key: string | null = "S01") {
        press.move(x, y);
        paint.move(x, y, key);
      },
      up() {
        press.up();
        const release = paint.up();
        // 브라우저는 손을 떼면 click 을 보낸다. 칠했으면 격자가 삼킨다.
        if (!release.swallowClick && press.click()) taps.push(1);
        return release;
      },
    };
  }

  it("약 0.5초 누르면 날짜 창이 한 번 열리고, 이어지는 click 은 탭(합격/불합격 전환)이 아니다", () => {
    const t = touch();
    t.down();
    vi.advanceTimersByTime(LONG_PRESS_MS - 1);
    expect(t.longs).toHaveLength(0);
    vi.advanceTimersByTime(1);
    expect(t.longs).toHaveLength(1);
    t.up();
    expect(t.taps).toHaveLength(0);
    // 다음 탭은 정상으로 돌아온다.
    t.down();
    vi.advanceTimersByTime(80);
    t.up();
    expect(t.taps).toHaveLength(1);
    expect(t.longs).toHaveLength(1);
  });

  it("짧게 누르고 떼면 탭이고 길게 누르기는 생기지 않는다", () => {
    const t = touch();
    t.down();
    vi.advanceTimersByTime(150);
    t.up();
    vi.advanceTimersByTime(LONG_PRESS_MS * 2);
    expect(t.taps).toHaveLength(1);
    expect(t.longs).toHaveLength(0);
  });

  it("손가락을 움직이면(드래그 합격) 길게 누르기가 취소되고 칠하기만 저장된다", () => {
    const t = touch();
    t.down(100, 100);
    vi.advanceTimersByTime(200);
    t.move(100 + MOVE_SLOP + 1, 100, "S01");
    t.move(180, 100, "S02");
    vi.advanceTimersByTime(LONG_PRESS_MS * 2);
    const release = t.up();
    expect(t.longs).toHaveLength(0);
    expect(t.taps).toHaveLength(0);
    expect(release.commit.map((h) => h.key)).toEqual(["S01", "S02"]);
  });

  it("세로로 움직이면(스크롤) 길게 누르기도 칠하기도 없다", () => {
    const t = touch();
    t.down(100, 100);
    t.move(100, 100 + MOVE_SLOP + 5);
    vi.advanceTimersByTime(LONG_PRESS_MS * 2);
    expect(t.longs).toHaveLength(0);
    expect(t.paint.phase).toBe("scroll");
  });

  it("손떨림(MOVE_SLOP 안)은 길게 누르기를 막지 않는다", () => {
    const t = touch();
    t.down(100, 100);
    t.move(100 + MOVE_SLOP, 100 - MOVE_SLOP);
    vi.advanceTimersByTime(LONG_PRESS_MS);
    expect(t.longs).toHaveLength(1);
  });
});
