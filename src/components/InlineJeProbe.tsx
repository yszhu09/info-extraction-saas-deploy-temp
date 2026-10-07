import {
  INFOWB_JE_VERSION,
  INFOWB_STATIC_CHUNKS_PATH,
} from "@/lib/jeDiagnostics";

type InlineJeProbeProps = {
  pageName: string;
};

function createProbeScript(pageName: string) {
  const chunksPath = JSON.stringify(INFOWB_STATIC_CHUNKS_PATH);
  const version = JSON.stringify(INFOWB_JE_VERSION);
  const page = JSON.stringify(pageName);

  return `(() => {
  const chunksPath = ${chunksPath};
  const version = ${version};
  const pageName = ${page};
  const selector = '[data-infowb-inline-probe-target="true"]';
  const write = (state, message) => {
    const element = document.querySelector(selector);
    if (!element) return;
    element.dataset.infowbProbeState = state;
    element.textContent = message;
  };
  const readyMessage = () =>
    "内联诊断通过：" + pageName + " React ready；安全静态资源前缀 " + chunksPath + " 可用。版本 " + version;
  const checkReady = () => {
    if (document.documentElement.dataset.infowbReactReady === "true") {
      write("react-ready", readyMessage());
      return true;
    }
    return false;
  };
  const boot = () => {
    document.documentElement.dataset.infowbInlineProbeLoaded = "true";
    if (checkReady()) return;
    window.setTimeout(() => {
      if (checkReady()) return;
      write(
        "react-timeout",
        "诊断提示：HTML 可见但 React 未就绪；请在 Network 查看 " + chunksPath + " 是否 403/blocked，并用页面版本 " + version + " 排查缓存。",
      );
    }, 5000);
  };
  window.addEventListener("infowb-react-ready", () => write("react-ready", readyMessage()));
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot, { once: true });
  } else {
    boot();
  }
})();`;
}

export function InlineJeProbe({ pageName }: InlineJeProbeProps) {
  return (
    <>
      <div
        data-testid="infowb-inline-probe"
        data-infowb-inline-probe-target="true"
        data-infowb-probe-state="server-rendered"
        className="border-b border-emerald-400/20 bg-slate-950 px-4 py-2 text-center text-xs font-medium text-emerald-100"
        suppressHydrationWarning
      >
        内联诊断已随 HTML 输出；等待浏览器脚本与 React ready。版本 {INFOWB_JE_VERSION}
      </div>
      <noscript>
        <div className="border-b border-red-300 bg-red-50 px-4 py-2 text-center text-sm font-semibold text-red-700">
          浏览器禁用了 JavaScript；信息提取/核对功能无法启动。请允许脚本并确认企业网络未拦截静态资源。
        </div>
      </noscript>
      <script
        id="infowb-je-inline-probe"
        data-infowb-version={INFOWB_JE_VERSION}
        dangerouslySetInnerHTML={{ __html: createProbeScript(pageName) }}
      />
    </>
  );
}
