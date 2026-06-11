/**
 * torture-coverage — 仕事文書 torture フィクスチャのカバレッジ回帰
 *
 * doc-translation-ir 同梱の work-doc-torture.docx（表/結合セル/混在言語・ハイパーリンク・
 * 脚注・追跡変更・段内太字）を reader にかけ、**全 probe が抽出され、文が分断されない**ことを
 * 検証する。v0.1（直下のみ走査）では 6/7 取りこぼし＋2文分断だったものが、再帰走査で 0 漏れに
 * なることを機械保証する常設テスト。
 *
 * 実行: tsx test/torture-coverage.ts
 */
import { readFileSync } from 'node:fs';
import { docxToDtir } from '../src/reader.js';
import { tortureDocxPath, tortureExpectedPath } from '@shuji-bonji/doc-translation-ir/fixtures';

const GREEN = (s: string) => `\x1b[32m${s}\x1b[0m`;
const RED = (s: string) => `\x1b[31m${s}\x1b[0m`;

interface Probe {
  key: string;
  text: string;
  where: string;
  expectExtractedNow: boolean;
  note?: string;
}

async function main(): Promise<void> {
  const buf = readFileSync(tortureDocxPath);
  const expected = JSON.parse(readFileSync(tortureExpectedPath, 'utf8')) as { probes: Probe[] };
  const dtir = await docxToDtir(buf, { fileName: 'work-doc-torture.docx', targetLang: 'en-GB' });

  const sources = dtir.segments.map((s) => s.text.source);
  const hasText = (needle: string) => sources.some((s) => s.includes(needle));

  const failures: string[] = [];
  const ok = (c: boolean, m: string) => {
    if (!c) failures.push(m);
  };

  // (1) 全 probe テキストが抽出されている（再帰走査の目標 = 0 漏れ）
  let extracted = 0;
  for (const pr of expected.probes) {
    const got = hasText(pr.text);
    if (got) extracted++;
    ok(got, `[probe:${pr.key}] "${pr.text}" 未抽出（${pr.where}）`);
  }

  // (2) 文が分断されていない（ハイパーリンク／追跡変更の前後ランが連結されている）
  ok(
    hasText('See the signed contract for the full terms.'),
    'ハイパーリンク段落が分断されている（前後ランと結合されていない）',
  );
  ok(
    hasText('The deadline is strictly binding.'),
    '追跡変更(w:ins)段落が分断されている（前後ランと結合されていない）',
  );

  // (3) role が構造を反映している
  const roleOf = (needle: string) =>
    dtir.segments.find((s) => s.text.source.includes(needle))?.role;
  ok(roleOf('Artikel 1') === 'table-cell', `表セルの role が table-cell でない（${roleOf('Artikel 1')}）`);
  ok(roleOf('Zusammenfassung der Bedingungen') === 'table-cell', '結合セルの role が table-cell でない');
  ok(
    roleOf('Vertrouwelijke voetnoot.') === 'footnote',
    `脚注の role が footnote でない（${roleOf('Vertrouwelijke voetnoot.')}）`,
  );

  // (4) 脚注パートが走査対象に入っている
  const parts = new Set(dtir.segments.map((s) => (s.anchor.ref as { part: string }).part));
  ok(parts.has('word/footnotes.xml'), 'footnotes.xml がパート列挙に入っていない');

  // (5) 表セルの混在言語が個別に解決されている（nl / fr / de）
  const langOf = (needle: string) =>
    dtir.segments.find((s) => s.text.source.includes(needle))?.language.value;
  ok(langOf('Artikel 1') === 'nl-NL', `表セル(nl)の言語: ${langOf('Artikel 1')}`);
  ok(langOf('Conditions générales') === 'fr-FR', `表セル(fr)の言語: ${langOf('Conditions générales')}`);
  ok(
    langOf('Zusammenfassung der Bedingungen') === 'de-DE',
    `結合セル(de)の言語: ${langOf('Zusammenfassung der Bedingungen')}`,
  );

  console.error(`coverage: ${extracted}/${expected.probes.length} probes 抽出`);
  console.error(`segments=${dtir.segments.length} translatable=${dtir.stats.translatableCount} groups=${dtir.stats.groupCount}`);
  console.error(`parts=${[...parts].join(', ')}`);
  console.error('');

  if (failures.length === 0) {
    console.error(GREEN(`TORTURE COVERAGE PASS — ${extracted}/${expected.probes.length} 漏れなし・文分断なし・role/言語/パート整合`));
    process.exit(0);
  }
  console.error(RED(`FAIL — ${failures.length} 件`));
  for (const f of failures) console.error('  - ' + f);
  process.exit(1);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
