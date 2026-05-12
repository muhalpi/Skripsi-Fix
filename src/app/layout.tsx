import type { Metadata } from "next";
import Script from "next/script";
import "./globals.css";

export const metadata: Metadata = {
  title: "Skripsi Helper Word Add-in",
  description: "Preset formatting, captions, and TOC helpers for Word skripsi documents.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>
        <Script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" strategy="beforeInteractive" />
        {children}
      </body>
    </html>
  );
}
