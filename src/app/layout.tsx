import type { Metadata } from "next";
import Script from "next/script";
import "./globals.css";

export const metadata: Metadata = {
  title: "Skripsi-Fix Word Add-in",
  description: "Preset formatting, captions, and TOC helpers for Word skripsi documents.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <head>
        <Script id="history-cache" strategy="beforeInteractive">
          {`
            window._historyCache = {
              replaceState: window.history && window.history.replaceState,
              pushState: window.history && window.history.pushState
            };
          `}
        </Script>
        <Script id="office-js" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" strategy="beforeInteractive" />
        <Script id="history-restore" strategy="beforeInteractive">
          {`
            (function keepHistoryApisAlive() {
              function restore() {
                if (!window._historyCache || !window.history) {
                  return;
                }
                if (window._historyCache.replaceState) {
                  window.history.replaceState = window._historyCache.replaceState;
                }
                if (window._historyCache.pushState) {
                  window.history.pushState = window._historyCache.pushState;
                }
              }

              restore();
              var attempts = 0;
              var timer = window.setInterval(function () {
                attempts += 1;
                restore();
                if (attempts >= 400) {
                  window.clearInterval(timer);
                }
              }, 25);
            })();
          `}
        </Script>
      </head>
      <body>{children}</body>
    </html>
  );
}
