/// <reference types="vite/client" />
/// <reference types="vite-plugin-pwa/client" />

interface ImportMetaEnv {
  readonly VITE_AZURE_CLIENT_ID?: string;
  readonly VITE_AZURE_AUTHORITY?: string;
  readonly VITE_AZURE_REDIRECT_URI?: string;
  /** vite.config 의 define 이 넣는 빌드 번호(짧은 커밋+시각) */
  readonly VITE_BUILD_ID?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
