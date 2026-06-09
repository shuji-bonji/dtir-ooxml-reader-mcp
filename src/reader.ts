/**
 * reader — docxToDtir: .docx (WordprocessingML) → DTIR
 *
 * 方針（DTIR 設計に忠実）:
 *  - reader は「セグメント表」だけ出す。書式XMLは anchor.ref に隠し、IR 表面に出さない。
 *  - id は anchor から決定的に導出（part＋XPath のハッシュ）。再実行で不変。
 *  - 翻訳対象は `<w:t>` のテキストのみ。複合フィールド・数値・空は translatable=false。
 *  - 言語はタグ×判定で調停（lang.ts）。
 *  - document.xml だけでなく header / footer パートも走査（取りこぼし防止）。
 */
import { createHash } from 'node:crypto';
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { isNumericOnly, resolveLanguage } from './lang.js';
import type {
  IRDocument,
  IRSegment,
  SegmentLanguage,
  SegmentRole,
  SegmentRun,
} from './dtir.js';

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

interface PartSpec {
  /** zip 内パス（例: word/document.xml）。 */
  part: string;
  /** ルート直下のパラグラフ容器タグ。 */
  container: 'w:body' | 'w:hdr' | 'w:ftr';
  /** 既定 role（header/footer は part が role を決める）。 */
  baseRole: SegmentRole | null;
}

export interface DocxToDtirOptions {
  fileName?: string;
  /** 翻訳先言語（DTIR language.target に入れる）。reader は翻訳しない。 */
  targetLang?: string | null;
}

// --- DOM ヘルパ -------------------------------------------------------------
type El = Element;

function childElements(el: El): El[] {
  const out: El[] = [];
  const nodes = el.childNodes;
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes.item(i);
    if (n && n.nodeType === ELEMENT_NODE) out.push(n as unknown as El);
  }
  return out;
}

function firstChildByTag(el: El, tag: string): El | null {
  for (const c of childElements(el)) if (c.tagName === tag) return c;
  return null;
}

function attr(el: El | null, name: string): string | null {
  return el ? el.getAttribute(name) : null;
}

/** w:t 要素の直下テキストを連結。 */
function textOfWt(wt: El): string {
  let s = '';
  const nodes = wt.childNodes;
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes.item(i);
    if (n && n.nodeType === TEXT_NODE) s += n.nodeValue ?? '';
  }
  return s;
}

// --- ラン抽出 ---------------------------------------------------------------
interface RunInfo {
  text: string;
  val: string | null; // run-level w:lang w:val（明示タグ）
  eastAsia: string | null;
  preserve: boolean;
}

function extractRun(r: El): RunInfo {
  const rPr = firstChildByTag(r, 'w:rPr');
  const lang = rPr ? firstChildByTag(rPr, 'w:lang') : null;
  let text = '';
  let preserve = false;
  for (const c of childElements(r)) {
    if (c.tagName === 'w:t') {
      text += textOfWt(c);
      if (attr(c, 'xml:space') === 'preserve') preserve = true;
    } else if (c.tagName === 'w:tab') {
      text += '\t';
    }
    // w:fldChar / w:instrText は翻訳対象テキストに含めない
  }
  return {
    text,
    val: attr(lang, 'w:val'),
    eastAsia: attr(lang, 'w:eastAsia'),
    preserve,
  };
}

function paragraphHasField(p: El): boolean {
  return p.getElementsByTagName('w:fldChar').length > 0;
}

function paragraphInstrText(p: El): string {
  const nodes = p.getElementsByTagName('w:instrText');
  let s = '';
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes.item(i);
    if (n) s += (n as unknown as El).textContent ?? '';
  }
  return s;
}

function paragraphStyle(p: El): string | null {
  const pPr = firstChildByTag(p, 'w:pPr');
  const pStyle = pPr ? firstChildByTag(pPr, 'w:pStyle') : null;
  return attr(pStyle, 'w:val');
}

// --- id 決定論導出 ----------------------------------------------------------
function deriveId(part: string, path: string): string {
  const h = createHash('sha1').update(`docx::${part}::${path}`).digest('hex');
  return `seg_${h.slice(0, 10)}`;
}

// --- container 既定言語 ------------------------------------------------------
function readContainerDefaultLang(stylesXml: string | null): string | null {
  if (!stylesXml) return null;
  const doc = new DOMParser().parseFromString(stylesXml, 'text/xml');
  const def = doc.getElementsByTagName('w:rPrDefault');
  if (def.length === 0) return null;
  const lang = (def.item(0) as unknown as El).getElementsByTagName('w:lang');
  if (lang.length === 0) return null;
  return (lang.item(0) as unknown as El).getAttribute('w:val');
}

// --- メイン -----------------------------------------------------------------
export async function docxToDtir(
  buf: Buffer | Uint8Array,
  options: DocxToDtirOptions = {},
): Promise<IRDocument> {
  const zip = await JSZip.loadAsync(buf);

  const read = async (p: string): Promise<string | null> => {
    const f = zip.file(p);
    return f ? f.async('string') : null;
  };

  const stylesXml = await read('word/styles.xml');
  const containerDefault = readContainerDefaultLang(stylesXml);

  // 走査対象パートを列挙（document → header* → footer*）
  const headerParts = Object.keys(zip.files)
    .filter((p) => /^word\/header\d*\.xml$/.test(p))
    .sort();
  const footerParts = Object.keys(zip.files)
    .filter((p) => /^word\/footer\d*\.xml$/.test(p))
    .sort();

  const specs: PartSpec[] = [
    { part: 'word/document.xml', container: 'w:body', baseRole: null },
    ...headerParts.map((p): PartSpec => ({ part: p, container: 'w:hdr', baseRole: 'header' })),
    ...footerParts.map((p): PartSpec => ({ part: p, container: 'w:ftr', baseRole: 'footer' })),
  ];

  const segments: IRSegment[] = [];
  let order = 0;

  for (const spec of specs) {
    const xml = await read(spec.part);
    if (!xml) continue;
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const containerEl = doc.getElementsByTagName(spec.container).item(0);
    if (!containerEl) continue;
    const paragraphs = childElements(containerEl as unknown as El).filter(
      (e) => e.tagName === 'w:p',
    );

    paragraphs.forEach((p, idx) => {
      const path = `/${spec.container}/w:p[${idx + 1}]`;
      const id = deriveId(spec.part, path);

      const isField = paragraphHasField(p);
      const runEls = childElements(p).filter((e) => e.tagName === 'w:r');
      const runs = runEls.map(extractRun).filter((r) => r.text.length > 0);
      const source = runs.map((r) => r.text).join('');
      const preserve = runs.some((r) => r.preserve);

      // role 決定
      let role: SegmentRole;
      if (spec.baseRole) {
        role = spec.baseRole;
      } else if (isField && /\bTOC\b/i.test(paragraphInstrText(p))) {
        role = 'toc';
      } else if ((paragraphStyle(p) ?? '').startsWith('Heading')) {
        role = 'heading';
      } else {
        role = 'body';
      }

      // 分類 + 言語
      let translatable: boolean;
      let skipReason: IRSegment['skipReason'];
      let language: SegmentLanguage;

      if (isField) {
        translatable = false;
        skipReason = 'field';
        language = { value: null, confidence: 0, source: 'default' };
      } else if (source.trim() === '') {
        translatable = false;
        skipReason = 'empty';
        language = { value: null, confidence: 0, source: 'default' };
      } else if (isNumericOnly(source)) {
        translatable = false;
        skipReason = 'numeric';
        language = { value: null, confidence: 0, source: 'detect' };
      } else {
        translatable = true;
        skipReason = null;
        language = resolveLanguage({
          source,
          explicitVals: runs.map((r) => r.val).filter((v): v is string => !!v),
          explicitEastAsia: runs.map((r) => r.eastAsia).filter((v): v is string => !!v),
          containerDefault,
        });
      }

      // text.runs（ラン分断時のみ）
      const textRuns: SegmentRun[] = [];
      if (runs.length > 1) {
        let pos = 0;
        runs.forEach((r, ri) => {
          const start = pos;
          const end = pos + r.text.length;
          textRuns.push({ runId: `${id}-r${ri}`, start, end });
          pos = end;
        });
      }
      const runIds = runs.map((_, ri) => `${id}-r${ri}`);

      const seg: IRSegment = {
        id,
        order: order++,
        anchor: { format: 'docx', ref: { part: spec.part, path, runIds } },
        role,
        text: {
          source,
          hasInlineFormatting: runs.length > 1,
          ...(textRuns.length > 0 ? { runs: textRuns } : {}),
          space: preserve ? 'preserve' : 'default',
        },
        language,
        translatable,
        skipReason,
        group: translatable ? language.value : null,
        context: { prev: null, next: null, parent: null }, // 後段でリンク
        translation: null,
        quality: null,
      };

      segments.push(seg);
    });
  }

  // context を読み順でリンク
  for (let i = 0; i < segments.length; i++) {
    segments[i].context.prev = i > 0 ? segments[i - 1].id : null;
    segments[i].context.next = i < segments.length - 1 ? segments[i + 1].id : null;
  }

  // 文書レベルの言語情報
  const presentLangs = [
    ...new Set(
      segments
        .filter((s) => s.translatable && s.language.value)
        .map((s) => s.language.value as string),
    ),
  ];
  const distinct = presentLangs.length;
  const isMultilingual = distinct > 1;

  const buffer = Buffer.isBuffer(buf) ? buf : Buffer.from(buf);
  const sha256 = createHash('sha256').update(buffer).digest('hex');

  const translatableCount = segments.filter((s) => s.translatable).length;
  const groupCount = new Set(
    segments.map((s) => s.group).filter((g): g is string => g !== null),
  ).size;

  return {
    irVersion: '0.1',
    source: {
      format: 'docx',
      fileName: options.fileName ?? 'document.docx',
      sha256,
      byteSize: buffer.length,
    },
    language: {
      default: {
        value: containerDefault,
        source: containerDefault ? 'container-default' : 'none',
      },
      target: options.targetLang ?? null,
      multilingual: {
        isMultilingual,
        score: isMultilingual ? Number(((distinct - 1) / distinct).toFixed(2)) : 0,
        method: 'per-segment',
        languagesPresent: presentLangs,
      },
    },
    segments,
    stats: {
      segmentCount: segments.length,
      translatableCount,
      groupCount,
    },
  };
}
