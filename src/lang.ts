/**
 * lang — 言語解決と分類
 *
 * DTIR の核心「言語は値ではなくヒント＋確信度＋出所」を実装する。
 * 優先順位:
 *   1. 明示的な run-level `<w:lang w:val>`（コンテナ既定の継承ではない）→ source=tag
 *   2. それが無ければローカル言語判定(franc) → source=detect
 *   3. 判定不能/短文 → コンテナ既定 → source=default
 *   4. 数値・記号のみ → 非言語、language=null（翻訳対象外）
 */
import { detectAll } from 'tinyld';
import type { LanguageCandidate, SegmentLanguage } from '@shuji-bonji/doc-translation-ir';

/**
 * 言語判定は tinyld（純JS）。franc は短文で誤判定が多い
 * （独語文を仏語に誤るなど）ため不採用。tinyld は ISO 639-1 と accuracy を返す。
 */

/** ISO 639-1 → BCP 47 の最小マップ（PoC 範囲）。 */
const ISO1_TO_BCP47: Record<string, string> = {
  nl: 'nl-NL',
  fr: 'fr-FR',
  de: 'de-DE',
  en: 'en-GB',
  ja: 'ja-JP',
  es: 'es-ES',
  it: 'it-IT',
  pt: 'pt-PT',
};

function toBcp47(iso1: string): string | null {
  return ISO1_TO_BCP47[iso1] ?? null;
}

/** これ未満の accuracy は「判定不能」とみなし既定にフォールバック。 */
const DETECT_THRESHOLD = 0.2;

/** 数値・記号のみ（文字を1つも含まず数字を含む）かどうか。 */
export function isNumericOnly(s: string): boolean {
  const hasLetter = /\p{L}/u.test(s);
  const hasDigit = /\p{Nd}/u.test(s);
  return !hasLetter && hasDigit;
}

/** source に含まれる Unicode スクリプト名の集合（混在検証・補助情報）。 */
export function detectScripts(s: string): string[] {
  const checks: [string, RegExp][] = [
    ['Latin', /\p{Script=Latin}/u],
    ['Han', /\p{Script=Han}/u],
    ['Hiragana', /\p{Script=Hiragana}/u],
    ['Katakana', /\p{Script=Katakana}/u],
    ['Hangul', /\p{Script=Hangul}/u],
    ['Cyrillic', /\p{Script=Cyrillic}/u],
    ['Arabic', /\p{Script=Arabic}/u],
    ['Hebrew', /\p{Script=Hebrew}/u],
  ];
  return checks.filter(([, re]) => re.test(s)).map(([name]) => name);
}

export interface ResolveInput {
  /** 連結後の原文。 */
  source: string;
  /** 明示的な run-level w:val（複数ランの値）。継承既定は含めない。 */
  explicitVals: string[];
  /** 明示的な w:eastAsia（あれば候補に加える）。 */
  explicitEastAsia: string[];
  /** コンテナ既定言語（styles docDefaults）。 */
  containerDefault: string | null;
}

/**
 * 言語を解決して SegmentLanguage を返す。
 * 翻訳対象テキスト専用（数値・フィールドは呼び出し側で除外済みの前提）。
 */
export function resolveLanguage(input: ResolveInput): SegmentLanguage {
  const { source, explicitVals, explicitEastAsia, containerDefault } = input;

  // 1. 明示タグ優先
  if (explicitVals.length > 0) {
    const value = mostCommon(explicitVals);
    const candidates: LanguageCandidate[] = [{ value, confidence: 0.95 }];
    // eastAsia は別スクリプトの候補として添える（混在スクリプトの手掛かり）
    for (const ea of dedupe(explicitEastAsia)) {
      if (ea !== value) candidates.push({ value: ea, confidence: 0.5 });
    }
    return { value, confidence: 0.95, source: 'tag', candidates };
  }

  // 2. ローカル判定（tinyld）— 採否に関わらず候補分布は計算しておき、
  //    フォールバック時も下流（言語解決ステージ）へ残す（捨てない）。
  const ranked = detectAll(source);
  const detectCandidates: LanguageCandidate[] = ranked
    .slice(0, 3)
    .map((r) => ({ value: toBcp47(r.lang) ?? r.lang, confidence: Number(r.accuracy.toFixed(2)) }));
  const top = ranked[0];
  if (top && top.accuracy >= DETECT_THRESHOLD) {
    const mapped = toBcp47(top.lang);
    if (mapped) {
      const confidence = Number(top.accuracy.toFixed(2));
      return { value: mapped, confidence, source: 'detect', candidates: detectCandidates };
    }
    // 閾値は超えたが 8言語マップ外。value は確定できないが、検出シグナルは下記で残す。
  }

  // 3. 既定にフォールバック。value/confidence/source は不変のまま、
  //    検出候補があれば添える（閾値未満・マップ外でも下流が分布を見られる）。
  const withCandidates = detectCandidates.length > 0 ? { candidates: detectCandidates } : {};
  if (containerDefault) {
    return { value: containerDefault, confidence: 0.3, source: 'default', ...withCandidates };
  }
  // 完全に検出不能（ranked 空）なら candidates は付けない＝「すべて失敗」は妥当に空。
  return { value: null, confidence: 0, source: 'default', ...withCandidates };
}

function mostCommon(arr: string[]): string {
  const counts = new Map<string, number>();
  for (const a of arr) counts.set(a, (counts.get(a) ?? 0) + 1);
  let best = arr[0];
  let bestN = 0;
  for (const [k, n] of counts) if (n > bestN) [best, bestN] = [k, n];
  return best;
}

function dedupe(arr: string[]): string[] {
  return [...new Set(arr)];
}
