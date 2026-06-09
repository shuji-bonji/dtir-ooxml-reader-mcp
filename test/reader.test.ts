/**
 * vitest spec — docxToDtir の主要不変条件。
 * 詳細な groundtruth 照合は test/run-fixture.ts（npm run test:fixture）。
 */
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { docxToDtir } from '../src/reader.js';
import { validateDtir } from '@shuji-bonji/doc-translation-ir/validate';
import { fixtureDocxPath as fixture } from '@shuji-bonji/doc-translation-ir/fixtures';

describe('docxToDtir', () => {
  it('produces a semantically valid multilingual DTIR', async () => {
    const dtir = await docxToDtir(readFileSync(fixture), {
      fileName: 'mixed-nl-fr-de-tricky.docx',
      targetLang: 'en-GB',
    });
    expect(validateDtir(dtir)).toEqual([]);
    expect(dtir.language.multilingual.isMultilingual).toBe(true);
    expect(dtir.language.default.value).toBe('nl-NL');
  });

  it('overrides an inherited default tag via detection (de paragraph w/o w:lang)', async () => {
    const dtir = await docxToDtir(readFileSync(fixture));
    const de = dtir.segments.find((s) => s.text.source.startsWith('Die Produktion'));
    expect(de?.language.value).toBe('de-DE');
    expect(de?.language.source).toBe('detect');
  });

  it('skips fields and numerics, keeps run-split offsets', async () => {
    const dtir = await docxToDtir(readFileSync(fixture));
    const fr = dtir.segments.find((s) => s.text.source.includes('prévisions'));
    expect(fr?.text.hasInlineFormatting).toBe(true);
    expect(fr?.text.runs?.length).toBe(3);
    const numeric = dtir.segments.find((s) => s.skipReason === 'numeric');
    expect(numeric?.translatable).toBe(false);
    expect(dtir.segments.some((s) => s.skipReason === 'field')).toBe(true);
  });

  it('is deterministic: ids stable across runs', async () => {
    const a = await docxToDtir(readFileSync(fixture));
    const b = await docxToDtir(readFileSync(fixture));
    expect(a.segments.map((s) => s.id)).toEqual(b.segments.map((s) => s.id));
  });
});
