export const metadata = {
  title: "msgraph-mcp-server",
  description:
    "Microsoft Graph API MCP Server — Exchange・Teams・OneDrive・SharePoint の 45 ツール",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
