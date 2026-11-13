/** 응답이 없어 기다리기를 그만뒀다는 오류. 로그인 문제가 아니라 '대기'로 남겨 다시 보낸다. */
export class TimeoutError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "TimeoutError";
  }
}

/** 동기화 묶음 전체를 멈춘(중단한) 오류. 기록은 실패가 아니라 대기로 남긴다. */
export const ABORTED_TEXT = "응답이 없어 보내기를 멈췄습니다. 기록은 대기로 남아 있으니 '다시 시도'를 누르세요.";

/** 토큰 받기가 시간 안에 끝나지 않음. */
export const TOKEN_TIMEOUT_TEXT = "Microsoft 토큰 응답이 없어(시간 초과) 멈췄습니다. 기록은 대기로 남습니다.";

export function isTimeoutLike(error: unknown): boolean {
  return error instanceof TimeoutError || (error instanceof Error && (error.name === "TimeoutError" || error.name === "AbortError"));
}

/**
 * promise 가 ms 안에 끝나지 않으면 TimeoutError. 원래 promise 는 취소할 수 없지만 더는 기다리지 않는다.
 * ms <= 0 이면 시간 제한 없음.
 */
export function withTimeout<T>(promise: Promise<T>, ms: number, message: string): Promise<T> {
  if (!(ms > 0) || !Number.isFinite(ms)) return promise;
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new TimeoutError(message)), ms);
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        clearTimeout(timer);
        reject(error);
      },
    );
  });
}
