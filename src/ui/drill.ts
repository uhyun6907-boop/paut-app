import { useEffect, useRef } from "react";

const MARK = "pautLayer";

type Layer = { id: number; close: () => void };

const stack: Layer[] = [];
let seq = 0;
let skipPops = 0;
let listening = false;

function ensureListener(): void {
  if (listening || typeof window === "undefined") return;
  listening = true;
  const state = history.state as Record<string, unknown> | null;
  if (!state?.pautRoot && state?.[MARK] == null) {
    history.replaceState({ ...(state ?? {}), pautRoot: true }, "");
  }
  window.addEventListener("popstate", () => {
    if (skipPops > 0) {
      skipPops -= 1;
      return;
    }
    stack.pop()?.close();
  });
}

function armLayer(close: () => void): number {
  ensureListener();
  const id = ++seq;
  history.pushState({ pautLayer: id }, "");
  stack.push({ id, close });
  return id;
}

function rebind(id: number, close: () => void): void {
  const layer = stack.find((item) => item.id === id);
  if (layer) layer.close = close;
}

function releaseLayer(id: number): void {
  const index = stack.findIndex((item) => item.id === id);
  if (index < 0) return;
  const top = index === stack.length - 1;
  stack.splice(index, 1);
  const state = history.state as Record<string, unknown> | null;
  if (top && state?.[MARK] === id) {
    skipPops += 1;
    history.back();
  }
}

/** 열린 동안 히스토리 한 칸을 넣어, 스와이프·뒤로가기가 앱을 나가지 않고 이 화면을 닫게 한다. */
export function useDrill(open: boolean, close: () => void): void {
  const closeRef = useRef(close);
  closeRef.current = close;
  const idRef = useRef<number | null>(null);
  const remount = useRef(false);

  useEffect(() => {
    if (!open) return;
    const bind = () => {
      idRef.current = null;
      closeRef.current();
    };
    if (remount.current && idRef.current != null) {
      remount.current = false;
      rebind(idRef.current, bind);
    } else {
      idRef.current = armLayer(bind);
    }
    const id = idRef.current;
    return () => {
      remount.current = true;
      queueMicrotask(() => {
        if (!remount.current) return;
        remount.current = false;
        if (idRef.current === id) idRef.current = null;
        if (id != null) releaseLayer(id);
      });
    };
  }, [open]);
}
