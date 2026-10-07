"use client";

import { useEffect, useState } from "react";
import {
  INFOWB_ASSET_PREFIX,
  INFOWB_JE_VERSION,
} from "@/lib/jeDiagnostics";

type JeStatusBadgeProps = {
  page: "extract" | "compare";
};

export function JeStatusBadge({ page }: JeStatusBadgeProps) {
  const [ready, setReady] = useState(false);

  useEffect(() => {
    document.documentElement.dataset.infowbReactReady = "true";
    document.documentElement.dataset.infowbPage = page;
    window.dispatchEvent(new Event("infowb-react-ready"));
    const readyFrame = window.requestAnimationFrame(() => setReady(true));

    return () => window.cancelAnimationFrame(readyFrame);
  }, [page]);

  return (
    <div
      data-testid="infowb-je-status"
      data-infowb-ready={ready ? "true" : "false"}
      className="flex flex-wrap items-center gap-2 rounded-2xl border border-emerald-300/25 bg-emerald-300/10 px-3 py-1.5 text-xs text-emerald-50"
    >
      <span>版本：{INFOWB_JE_VERSION}</span>
      <span
        className={
          ready
            ? "rounded-full bg-emerald-300 px-2 py-0.5 font-semibold text-emerald-950"
            : "rounded-full bg-amber-200 px-2 py-0.5 font-semibold text-amber-950"
        }
      >
        {ready ? "● 浏览器功能已就绪" : "● 等待浏览器功能就绪"}
      </span>
      <span className="text-emerald-100/80">
        静态资源：{INFOWB_ASSET_PREFIX}/_next/static
      </span>
    </div>
  );
}
