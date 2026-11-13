import { describe, expect, it } from "vitest";
import { PAGES_REDIRECT_URI, buildRedirectUri } from "./msal";

describe("MSAL redirect URI", () => {
  it("Pages 빌드는 등록된 주소와 끝 슬래시가 같다", () => {
    expect(buildRedirectUri("/paut-app/", "https://uhyun6907-boop.github.io")).toBe(PAGES_REDIRECT_URI);
    expect(buildRedirectUri("/paut-app/", "https://example.com")).toBe("https://uhyun6907-boop.github.io/paut-app/");
  });

  it("루트 빌드는 그 출처의 루트다", () => {
    expect(buildRedirectUri("/", "https://example.com")).toBe("https://example.com/");
  });
});