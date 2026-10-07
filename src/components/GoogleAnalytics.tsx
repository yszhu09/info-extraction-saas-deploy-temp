import Script from "next/script";

/** GA4 Measurement ID for infoworkbench.online (property 557896927). */
const GA_MEASUREMENT_ID = "G-HDYFFTSVJZ";

/**
 * Injects GA4 gtag. Uses NEXT_PUBLIC_GA_ID when set, else the committed ID
 * (same pattern as sibling sites that hardcode G-…).
 * Redeploy touch: 2026-10-07 production restore.
 */
export function GoogleAnalytics() {
  const gaId = (process.env.NEXT_PUBLIC_GA_ID?.trim() || GA_MEASUREMENT_ID).trim();
  if (!gaId) return null;

  const inline = `
window.dataLayer = window.dataLayer || [];
function gtag(){dataLayer.push(arguments);}
gtag('js', new Date());
gtag('config', '${gaId}');
`;

  return (
    <>
      <Script
        src={`https://www.googletagmanager.com/gtag/js?id=${gaId}`}
        strategy="afterInteractive"
      />
      <Script id="ga4-gtag" strategy="afterInteractive">
        {inline}
      </Script>
    </>
  );
}
