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
} from '@shuji-bonji/doc-translation-ir';

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

interface PartSpec {
  /** zip 内パス（例: word/document.xml）。 */
  part: string;
  /** 既定 role（header/footer/footnote/endnote は part が role を決める）。 */
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

// --- 再帰走査（v0.2） -------------------------------------------------------
// reader は「<w:t> を持つ段落(w:p)」を、容器直下に限らず**どの入れ子からも**拾う。
// 表(w:tbl/w:tr/w:tc)・SDT・テキストボックス内の段落、脚注/文末脚注パートを網羅する。

interface ParaHit {
  /** 対象段落。 */
  p: El;
  /** part ルート（documentElement）からの構造パス。例 /w:body[1]/w:tbl[1]/w:tr[2]/w:tc[1]/w:p[1] */
  path: string;
  /** 表セル内か（role=table-cell の判定に使う）。 */
  inTableCell: boolean;
}

/**
 * documentElement 起点で w:p を再帰収集する。各要素には**同名兄弟内の1始まり連番**を
 * 付け、writer 側の汎用ナビゲータと一致するパスを作る。
 * w:p には降りない（段落は入れ子にならない）。脚注の separator/continuationSeparator
 * （w:type 付き）は本文ではないので走査しない。
 */
function collectParagraphs(el: El, prefix: string, inTableCell: boolean, hits: ParaHit[]): void {
  const counts = new Map<string, number>();
  for (const c of childElements(el)) {
    const tag = c.tagName;
    const idx = (counts.get(tag) ?? 0) + 1;
    counts.set(tag, idx);
    const childPath = `${prefix}/${tag}[${idx}]`;
    if (tag === 'w:p') {
      hits.push({ p: c, path: childPath, inTableCell });
      continue; // 段落内には w:p は無い
    }
    if ((tag === 'w:footnote' || tag === 'w:endnote') && c.getAttribute('w:type')) {
      continue; // separator / continuationSeparator は本文でない
    }
    collectParagraphs(c, childPath, inTableCell || tag === 'w:tc', hits);
  }
}

/**
 * 段落内の**テキストを持つラン**を読み順で収集（再帰）。
 * w:hyperlink / w:ins / w:smartTag / w:sdt 内の run も連結対象＝文が分断されない。
 * w:del（削除済みテキスト）配下は対象外。run は入れ子にならないので降りない。
 */
function collectRunEls(p: El): El[] {
  const out: El[] = [];
  const walk = (el: El): void => {
    for (const c of childElements(el)) {
      if (c.tagName === 'w:del') continue; // 削除済みテキストは翻訳対象外
      if (c.tagName === 'w:r') {
        out.push(c);
        continue;
      }
      walk(c);
    }
  };
  walk(p);
  return out;
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

  const hasFootnotes = !!zip.file('word/footnotes.xml');
  const hasEndnotes = !!zip.file('word/endnotes.xml');

  const specs: PartSpec[] = [
    { part: 'word/document.xml', baseRole: null },
    ...headerParts.map((p): PartSpec => ({ part: p, baseRole: 'header' })),
    ...footerParts.map((p): PartSpec => ({ part: p, baseRole: 'footer' })),
    ...(hasFootnotes ? [{ part: 'word/footnotes.xml', baseRole: 'footnote' } as PartSpec] : []),
    ...(hasEndnotes ? [{ part: 'word/endnotes.xml', baseRole: 'endnote' } as PartSpec] : []),
  ];

  const segments: IRSegment[] = [];
  let order = 0;

  for (const spec of specs) {
    const xml = await read(spec.part);
    if (!xml) continue;
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const root = doc.documentElement as unknown as El | null;
    if (!root) continue;
    const hits: ParaHit[] = [];
    collectParagraphs(root, '', false, hits);

    hits.forEach((hit) => {
      const p = hit.p;
      const path = hit.path;
      const id = deriveId(spec.part, path);

      const isField = paragraphHasField(p);
      const runEls = collectRunEls(p);
      const runs = runEls.map(extractRun).filter((r) => r.text.length > 0);
      const source = runs.map((r) => r.text).join('');
      const preserve = runs.some((r) => r.preserve);

      // role 決定（part 既定 > 表セル > TOC > 見出し > body）
      let role: SegmentRole;
      if (spec.baseRole) {
        role = spec.baseRole;
      } else if (hit.inTableCell) {
        role = 'table-cell';
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
