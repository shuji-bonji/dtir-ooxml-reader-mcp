#!/usr/bin/env node
/**
 * dtir-ooxml-reader-mcp MCP server
 *
 * tool: docx_to_dtir — base64 の .docx を DTIR(セグメント表) に変換して返す。
 * reader は翻訳しない。後段の言語解決/翻訳/writer は別 MCP が担う。
 *
 * 鉄則（shuji-mcp-patterns）: stdout を汚さない。ログは console.error のみ。
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { docxToDtir } from './reader.js';

const server = new McpServer({ name: 'dtir-ooxml-reader-mcp', version: '0.0.1' });

server.tool(
  'docx_to_dtir',
  '.docx (WordprocessingML) を DTIR セグメント表に変換する。混在言語ドキュメントの ' +
    '翻訳パイプラインの入口。<w:t> のみを翻訳対象とし、複合フィールド・数値・空は ' +
    'translatable=false。言語はタグ×ローカル判定で調停。header/footer も走査する。',
  {
    docxBase64: z.string().describe('base64 エンコードした .docx バイナリ'),
    fileName: z.string().optional().describe('元ファイル名（メタ情報）'),
    targetLang: z
      .string()
      .optional()
      .describe('翻訳先言語 BCP47（DTIR language.target に格納。reader は翻訳しない）'),
  },
  async (args) => {
    try {
      const buf = Buffer.from(args.docxBase64, 'base64');
      const dtir = await docxToDtir(buf, {
        fileName: args.fileName,
        targetLang: args.targetLang ?? null,
      });
      return { content: [{ type: 'text', text: JSON.stringify(dtir, null, 2) }] };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return {
        isError: true,
        content: [{ type: 'text', text: `docx_to_dtir failed: ${msg}` }],
      };
    }
  },
);

const transport = new StdioServerTransport();
await server.connect(transport);
console.error('dtir-ooxml-reader-mcp MCP server running on stdio');
