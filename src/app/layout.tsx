import type { Metadata } from "next";

import "./globals.css";
import { Providers } from "./providers";

export const metadata: Metadata = {
  title: "Axilus CRC Operations",
  description: "Centre de relations clients — tableau de bord opérationnel Axilus",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="fr" suppressHydrationWarning>
      <body className="min-h-screen font-sans antialiased flex flex-col">
        <Providers>{children}</Providers>
      </body>
    </html>
  );
}
