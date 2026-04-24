import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "PnP Action Store Report",
  description: "PnP OOS, phantom stock & missing SKU reports — iRam / OuterJoin",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased">
        {children}
      </body>
    </html>
  );
}
