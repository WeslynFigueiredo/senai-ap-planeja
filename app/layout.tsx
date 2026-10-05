import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "SENAI AP Planeja",
  description: "Do plano ao Excel institucional. Ferramenta de apoio ao planejamento docente."
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="pt-BR">
      <body>{children}</body>
    </html>
  );
}
