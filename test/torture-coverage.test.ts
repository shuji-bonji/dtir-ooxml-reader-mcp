/**
 * torture-coverage spec — 仕事文書 torture フィクスチャの抽出カバレッジ回帰（vitest）
 *
 * validateDtir は「DTIR が契約として整合しているか」しか見ない。取りこぼし
 * （reader が表セルや脚注を丸ごと落とす）は、落とした後の DTIR が内部的に
 * 整合しているため validate では原理的に捕まらない。そこを埋めるのが本 spec。
 *
 * 設計＝ラチェット:
 *  doc-translation-ir 同梱の expected.json を唯一の真実源とし、各 probe について
 *  「実際に抽出されたか == expectExtractedNow」を assert する。
 *   - reader が退行（抽出していた probe を落とす）→ true!=false で fail
 *   - reader が改善（落としていた probe を拾う）→ false!=true で fail
 *     → expected.json のフラグ更新を強制し、カバレッジ変化を必ずレビューに乗せる
 *
 * 標準の `npm test`(vitest run) に載る常設テスト。なお test/torture-coverage.ts は
 * 色付きサマリを出す CLI 版（tsx 実行）で、内容は本 spec と等価。
 */
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { docxToDtir } from '../src/reader.js';
import type { IRDocument } from '@shuji-bonji/doc-translation-ir';
import {
  tortureDocxPath,
  tortureExpectedPath,
} from '@shuji-bonji/doc-translation-ir/fixtures';

interface Probe {
  key: string;
  text: string;
  where: string;
  expectExtractedNow: boolean;
  note?: string;
}
interface Expected {
  fixture: string;
  description: string;
  probes: Probe[];
}

const expected = JSON.parse(readFileSync(tortureExpectedPath, 'utf8')) as Expected;

let dtir: IRDocument;
let sources: string[];
const has = (needle: string): boolean => sources.some((s) => s.includes(needle));
const find = (needle: string) =>
  dtir.segments.find((s) => s.text.source.includes(needle));

beforeAll(async () => {
  dtir = await docxToDtir(readFileSync(tortureDocxPath), {
    fileName: 'work-doc-torture.docx',
    targetLang: 'en-GB',
  });
  sources = dtir.segments.map((s) => s.text.source);
});

describe('torture coverage — 抽出ラチェット', () => {
  it.each(expected.probes)(
    'probe $key ($where): 抽出 == expectExtractedNow',
    (p) => {
      expect(has(p.text)).toBe(p.expectExtractedNow);
    },
  );
});

describe('torture coverage — 文分断なし', () => {
  it('w:hyperlink 段落が前後ランと結合される', () => {
    expect(has('See the signed contract for the full terms.')).toBe(true);
  });
  it('w:ins(追跡変更) 段落が前後ランと結合される', () => {
    expect(has('The deadline is strictly binding.')).toBe(true);
  });
});

describe('torture coverage — role/言語/パート整合', () => {
  it('表セル/結合セルの role=table-cell・脚注=footnote', () => {
    expect(find('Artikel 1')?.role).toBe('table-cell');
    expect(find('Zusammenfassung der Bedingungen')?.role).toBe('table-cell');
    expect(find('Vertrouwelijke voetnoot.')?.role).toBe('footnote');
  });
  it('表セルの混在言語が個別解決される (nl/fr/de)', () => {
    expect(find('Artikel 1')?.language.value).toBe('nl-NL');
    expect(find('Conditions générales')?.language.value).toBe('fr-FR');
    expect(find('Zusammenfassung der Bedingungen')?.language.value).toBe('de-DE');
  });
  it('footnotes.xml がパート列挙に含まれる', () => {
    const parts = new Set(
      dtir.segments.map((s) => (s.anchor.ref as { part: string }).part),
    );
    expect(parts.has('word/footnotes.xml')).toBe(true);
  });
});
