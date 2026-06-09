/**
 * run-fixture — dtir-ooxml-reader-mcp の受け入れテスト
 *
 * doc-translation-ir のフィクスチャ docx を読み、reader 出力 DTIR を
 *   (1) JSON Schema（構造）
 *   (2) validate-dtir（意味整合）
 *   (3) groundtruth（実装非依存の期待値）
 * の3段で検証し、さらに id 決定性（2回実行で id 不変）を確認する。
 *
 * 実行: tsx test/run-fixture.ts
 */
import { readFileSync } from 'node:fs';
import Ajv2020 from 'ajv/dist/2020.js';
import addFormats from 'ajv-formats';
import { docxToDtir } from '../src/reader.js';
import { detectScripts } from '../src/lang.js';
import { validateDtir } from '@shuji-bonji/doc-translation-ir/validate';
import {
  fixtureDocxPath,
  fixtureGroundtruthPath,
  schemaPath,
} from '@shuji-bonji/doc-translation-ir/fixtures';

const fixtureDocx = fixtureDocxPath;
const groundtruthPath = fixtureGroundtruthPath;

const RED = (s: string) => `\x1b[31m${s}\x1b[0m`;
const GREEN = (s: string) => `\x1b[32m${s}\x1b[0m`;

interface GtSegment {
  key: string;
  part: string;
  role: string;
  expectSource: string | null;
  expectLang: string | null;
  expectLangSource: string;
  translatable: boolean;
  skipReason: string | null;
  runCount?: number;
  scripts?: string[];
}

async function main(): Promise<void> {
  const buf = readFileSync(fixtureDocx);
  const gt = JSON.parse(readFileSync(groundtruthPath, 'utf8')) as {
    containerDefaultLang: string;
    expectedLanguagesPresent: string[];
    segments: GtSegment[];
  };
  const schema = JSON.parse(readFileSync(schemaPath, 'utf8'));

  const dtir = await docxToDtir(buf, {
    fileName: 'mixed-nl-fr-de-tricky.docx',
    targetLang: 'en-GB',
  });

  const failures: string[] = [];
  const ok = (cond: boolean, msg: string) => {
    if (!cond) failures.push(msg);
  };

  // (1) JSON Schema
  const ajv = new Ajv2020({ allErrors: true, strict: false });
  addFormats(ajv);
  const validate = ajv.compile(schema);
  if (!validate(dtir)) {
    for (const e of validate.errors ?? []) {
      failures.push(`[schema] ${e.instancePath} ${e.message}`);
    }
  } else {
    console.error(GREEN('[1/4] JSON Schema: PASS'));
  }

  // (2) 意味整合
  const issues = validateDtir(dtir);
  if (issues.length) {
    for (const i of issues) failures.push(`[semantic] ${i.code} ${i.segmentId ?? '(doc)'}: ${i.message}`);
  } else {
    console.error(GREEN('[2/4] validate-dtir: PASS'));
  }

  // (3) groundtruth
  ok(
    dtir.language.default.value === gt.containerDefaultLang,
    `container default lang: got ${dtir.language.default.value} want ${gt.containerDefaultLang}`,
  );
  ok(dtir.language.multilingual.isMultilingual === true, 'should be multilingual');
  for (const want of gt.expectedLanguagesPresent) {
    ok(
      dtir.language.multilingual.languagesPresent.includes(want),
      `languagesPresent missing ${want} (got ${dtir.language.multilingual.languagesPresent.join(',')})`,
    );
  }

  for (const g of gt.segments) {
    const inPart = dtir.segments.filter(
      (s) => (s.anchor.ref as { part: string }).part === g.part,
    );
    let seg = undefined as (typeof dtir.segments)[number] | undefined;
    if (g.expectSource) {
      const want = g.expectSource.trim();
      seg = inPart.find(
        (s) => s.text.source.includes(want) || want.includes(s.text.source.trim()),
      );
    } else {
      seg = inPart.find((s) => s.role === g.role && s.skipReason === g.skipReason);
    }
    if (!seg) {
      failures.push(`[gt:${g.key}] セグメント未検出 (part=${g.part})`);
      continue;
    }
    ok(seg.role === g.role, `[gt:${g.key}] role got ${seg.role} want ${g.role}`);
    ok(seg.translatable === g.translatable, `[gt:${g.key}] translatable got ${seg.translatable} want ${g.translatable}`);
    ok(seg.skipReason === g.skipReason, `[gt:${g.key}] skipReason got ${seg.skipReason} want ${g.skipReason}`);
    ok(seg.language.value === g.expectLang, `[gt:${g.key}] lang got ${seg.language.value} want ${g.expectLang}`);
    ok(seg.language.source === g.expectLangSource, `[gt:${g.key}] langSource got ${seg.language.source} want ${g.expectLangSource}`);
    if (g.runCount !== undefined) {
      const runIds = (seg.anchor.ref as { runIds: string[] }).runIds;
      ok(runIds.length === g.runCount, `[gt:${g.key}] runCount got ${runIds.length} want ${g.runCount}`);
    }
    if (g.scripts) {
      const got = detectScripts(seg.text.source);
      for (const sc of g.scripts) {
        ok(got.includes(sc), `[gt:${g.key}] script ${sc} not detected (got ${got.join(',')})`);
      }
    }
  }
  if (failures.filter((f) => f.startsWith('[gt')).length === 0)
    console.error(GREEN('[3/4] groundtruth: PASS'));

  // (4) id 決定性
  const dtir2 = await docxToDtir(buf, { fileName: 'mixed-nl-fr-de-tricky.docx' });
  const ids1 = dtir.segments.map((s) => s.id).join(',');
  const ids2 = dtir2.segments.map((s) => s.id).join(',');
  if (ids1 !== ids2) failures.push('[determinism] id が再実行で不変でない');
  else console.error(GREEN('[4/4] id determinism: PASS'));

  console.error('');
  if (failures.length === 0) {
    console.error(GREEN(`ALL PASS — ${dtir.segments.length} segments, ${dtir.stats.translatableCount} translatable, ${dtir.stats.groupCount} groups`));
    process.exit(0);
  } else {
    console.error(RED(`FAIL — ${failures.length} 件`));
    for (const f of failures) console.error('  - ' + f);
    process.exit(1);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
