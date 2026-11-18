// Parity tests for the browser core in web/index.html against the Python reference (src/masume).
//
//   node web/test_core.mjs            # fixtures + fill parity + 300 random workbooks
//   FUZZ=2000 node web/test_core.mjs  # more random workbooks
//   PYTHON=python3 node web/test_core.mjs
//
// The <script id="core"> block of index.html is evaluated as-is (no build step).
import { readFileSync, writeFileSync, mkdtempSync, readdirSync, rmSync, existsSync, mkdirSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import os from 'node:os';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.dirname(HERE);
const SRC = path.join(ROOT, 'src');
const PY = process.env.PYTHON || 'python';
const FUZZ = Number(process.env.FUZZ ?? 300);
const KEEP = !!process.env.KEEP;
const TODAY = { y: 2026, m: 11, d: 19 };

const html = readFileSync(path.join(HERE, 'index.html'), 'utf8');
const core = /<script id="core">([\s\S]*?)<\/script>/.exec(html)[1];
vm.runInThisContext(core, { filename: 'index.html#core' });
const M = globalThis.Masume;

const TMP = mkdtempSync(path.join(os.tmpdir(), 'masume-web-'));
let failures = 0, checks = 0;
const fail = (msg) => { failures++; console.log('  FAIL ' + msg); };
const check = (cond, msg) => { checks++; if (!cond) fail(msg); return cond; };
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
function diffLine(a, b) {
  const sa = JSON.stringify(a, null, 1).split('\n'), sb = JSON.stringify(b, null, 1).split('\n');
  for (let i = 0; i < Math.max(sa.length, sb.length); i++) if (sa[i] !== sb[i]) return `line ${i}: js=${sa[i]} | py=${sb[i]}`;
  return '(no line diff)';
}
const pyEnv = { ...process.env, PYTHONPATH: SRC, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' };
const py = (args, input) => execFileSync(PY, args, { input, env: pyEnv, cwd: ROOT, maxBuffer: 1 << 28 }).toString('utf8').replace(/\r\n/g, '\n');

// ---- batch helper (one Python process for many files) -----------------------
const HELPER = path.join(TMP, 'helper.py');
writeFileSync(HELPER, String.raw`
import sys, json, zipfile, datetime as dt
sys.path.insert(0, sys.argv[1])
from masume import ooxml
from masume.detect import detect
from masume.render import workbook_to_markdown
from masume.cli import template_for
from masume.fill import fill
cmd = sys.argv[2]
jobs = json.loads(sys.stdin.buffer.read().decode("utf-8"))
TODAY = dt.date(2026, 11, 19)
out = []
for j in jobs:
    try:
        if cmd == "inspect":
            wb = ooxml.load(j)
            fs = detect(wb)
            out.append({"fields": [x.to_dict() for x in fs], "md": workbook_to_markdown(wb), "template": template_for(fs)})
        elif cmd == "fill":
            rep = fill(j["src"], j["dst"], j["data"], mark=j.get("mark", "■"), today=TODAY)
            out.append({"report": rep.to_dict()})
        elif cmd == "compare":
            diffs = []
            with zipfile.ZipFile(j["a"]) as a, zipfile.ZipFile(j["b"]) as b:
                if b.testzip() is not None:
                    diffs.append("testzip")
                if a.namelist() != b.namelist():
                    diffs.append("namelist")
                else:
                    for n in a.namelist():
                        if a.read(n) != b.read(n):
                            diffs.append(n)
            out.append(diffs)
        elif cmd == "values":
            wb = ooxml.load(j["file"])
            vals = []
            for ref in j["cells"]:
                s, r = ref.rsplit("!", 1)
                vals.append(wb.sheet(s).value(*ooxml.parse_ref(r)))
            out.append(vals)
    except Exception as e:
        out.append({"exception": type(e).__name__ + ": " + str(e)})
sys.stdout.buffer.write(json.dumps(out, ensure_ascii=False, default=str).encode("utf-8"))
`);
const helper = (cmd, jobs) => JSON.parse(execFileSync(PY, [HELPER, SRC, cmd], { input: JSON.stringify(jobs), env: pyEnv, maxBuffer: 1 << 28 }).toString('utf8'));

async function jsInspect(file) {
  const wb = await M.load(readFileSync(file));
  const fields = M.detect(wb);
  return { fields: fields.map((f) => f.toDict()), md: M.workbookToMarkdown(wb), template: M.templateFor(fields) };
}
async function jsFill(src, dst, data, mark = '■') {
  try {
    const wb = await M.load(readFileSync(src));
    const { report, bytes } = await M.fill(wb, data, { mark, today: TODAY });
    if (bytes && !report.errors.length) writeFileSync(dst, bytes);
    return { report: JSON.parse(JSON.stringify(report)) };
  } catch (e) {
    if (e instanceof M.FillError || e instanceof M.WriteError) throw e;
    return { exception: `${e.constructor.name}: ${e.message}` };
  }
}
function normReport(r) {
  // error messages are localised in the JS port: compare their count and the key they start with
  const errs = r.errors.map((e) => (e.startsWith('verify failed') ? 'verify' : e.split(': ')[0]));
  return { ...r, errors: errs };
}

// ============================================================================
console.log('1. formatter unit checks (mirrors tests/test_masume.py, tests/test_hard.py)');
{
  const eq = (a, b, msg) => check(same(a, b), `${msg}: got ${JSON.stringify(a)}, want ${JSON.stringify(b)}`);
  eq(M.formatDate('2019-05-01', '令和　　年　　月　　日', true), '令和元年5月1日', 'wareki 元年');
  eq(M.formatDate('2019-04-30', '令和　　年　　月　　日', true), '平成31年4月30日', 'wareki 平成31');
  eq(M.formatDate('2026/11/19', '　　年　　月　　日', false), '2026年11月19日', 'seireki');
  eq(M.formatDate('2026-11-19', '西暦　　年　　月　　日', true), '西暦2026年11月19日', '西暦');
  eq(M.formatDate('2026-11-19', '提出日：令和　年　月　日', true), '提出日：令和8年11月19日', 'pre text');
  eq(M.formatDate('未定', null, true), '未定', 'non-date');
  eq(M.formatDate('1990-04-02', '昭和・平成　　年　　月　　日生（満　　歳）', true, TODAY), '平成2年4月2日生（満36歳）', 'era list + age');
  eq(M.formatDate('1985-06-15', '明・大・昭・平・令　　年　　月　　日', true), '昭和60年6月15日', 'era list prints the full era name');
  for (const [inp, want] of [['昭和60年6月15日', [1985, 6, 15]], ['S60.6.15', [1985, 6, 15]], ['平成元年1月8日', [1989, 1, 8]],
    ['R8/11/19', [2026, 11, 19]], ['２０２６／１１／１９', [2026, 11, 19]], ['来週', null]]) {
    const d = M.parseDate(inp);
    eq(d && [d.y, d.m, d.d], want, `parseDate(${inp})`);
  }
  let bad = false;
  try { M.formatDate('来週', '令和　年　月　日', true); } catch (e) { bad = e instanceof M.FillError; }
  check(bad, 'unparseable date into a placeholder raises FillError');
  eq(M.formatPattern('　　時　　分', '09:30'), '9時30分', 'hour without leading zero');
  eq(M.formatChoice('□有　□無', '無'), '□有　■無', 'choice');
  eq(M.formatChoice('■有　□無', ['無'], '☑'), '□有　☑無', 'choice mark');
  eq(M.formatChoice('□本人　□同一世帯の方　□その他（　　）', { その他: '祖父' }), '□本人　□同一世帯の方　■その他（祖父）', 'choice extra');
  let threw = false;
  try { M.formatChoice('□有　□無', '多分'); } catch (e) { threw = e instanceof M.FillError; }
  check(threw, 'unknown option raises FillError');
  eq(M.formatPattern('〒　　　－', '100-0001'), '〒100－0001', 'pattern zip');
  eq(M.formatPattern('（　　　　）　　　　－', '03-1234-5678'), '（03）1234－5678', 'pattern tel');
  eq(M.formatPattern('　　時　　分', '9:05'), '9時05分', 'pattern time');
  // xml surgery
  const xml = '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<x:sheetData><x:row r="1" spans="1:1"><x:c r="A1" s="3"/></x:row>' +
    '<x:row r="5"><x:c r="C5"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>';
  const out = M.setCellsXml(xml, [[1, 1, 'あ<&>'], [1, 3, 7], [3, 2, 'new'], [5, 1, 'z']]);
  check(out.includes('<x:row r="1"><x:c r="A1" s="3" t="inlineStr"><x:is><x:t xml:space="preserve">あ&lt;&amp;&gt;</x:t></x:is></x:c><x:c r="C1"><x:v>7</x:v></x:c></x:row><x:row r="3">'), 'prefixed surgery + insert');
  check(!out.includes('spans='), 'spans removed');
  let refused = false;
  try { M.setCellsXml('<worksheet><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>', [[1, 1, 3]]); } catch (e) { refused = e instanceof M.WriteError; }
  check(refused, 'formula refused');
  const si = M.parseXML('<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><t>山田</t><rPh sb="0" eb="2"><t>ヤマダ</t></rPh><phoneticPr fontId="1"/></si>');
  eq(M.textOfSi(si), '山田', 'furigana ignored');
}

// ============================================================================
console.log('2. detection parity with `python -m masume inspect --json` (+ md, template)');
const fixtureDirs = ['tests/fixtures', 'tests/fixtures_hard', 'tests/fixtures_review'];
const fixtures = fixtureDirs.flatMap((d) => readdirSync(path.join(ROOT, d)).filter((f) => f.endsWith('.xlsx')).map((f) => path.join(ROOT, d, f)));
for (const file of fixtures) {
  const rel = path.relative(ROOT, file).replace(/\\/g, '/');
  const pyFields = JSON.parse(py(['-m', 'masume', 'inspect', '--json', file]));
  const pyMd = py(['-m', 'masume', 'md', file]);
  const pyTpl = JSON.parse(py(['-m', 'masume', 'template', file]));
  const js = await jsInspect(file);
  const ok1 = check(same(js.fields, pyFields), `${rel}: fields differ — ${diffLine(js.fields, pyFields)}`);
  const ok2 = check(js.md + '\n' === pyMd, `${rel}: markdown differs — ${diffLine((js.md + '\n').split('\n'), pyMd.split('\n'))}`);
  const ok3 = check(same(js.template, pyTpl), `${rel}: template differs`);
  console.log(`  ${ok1 && ok2 && ok3 ? 'ok  ' : 'DIFF'} ${rel} (${js.fields.length} fields: ${[...new Set(js.fields.map((f) => f.kind))].join(', ')})`);
}

// ============================================================================
console.log('3. fill parity (same report, byte-identical parts, values read back by the Python loader)');
const FX = (p) => path.join(ROOT, p);
const fillCases = [
  // tests/test_masume.py
  { src: 'tests/fixtures/kyuka_todoke.xlsx', data: { 日付: '2026-11-19', 氏名: '山田 太郎', 社員番号: 42, 休暇の種類: '代休', 開始日: '2026-12-01', 備考: 'なし' },
    expect: { '休暇届!T5': '令和8年11月19日', '休暇届!H9': '山田 太郎', '休暇届!X8': 42, '休暇届!H12': '□年次有給休暇　□特別休暇　■代休　□その他', '休暇届!H13': '令和8年12月1日', '休暇届!B21': '備考：なし' } },
  { src: 'tests/fixtures/excel_saved_kyuka.xlsx', data: { 氏名: '山田', 日付: '2026-11-19' }, expect: { '休暇届!H9': '山田' } },
  { src: 'tests/fixtures/keihi_seisan.xlsx', data: { 明細: [{ 日付: '11/1', 内容: '交通費', 金額: 2300 }, ['', '11/2', '書籍', 1800]], 支払方法: '現金' },
    expect: { '経費精算書!J10': '交通費', '経費精算書!Y10': 2300, '経費精算書!Y11': 1800, '経費精算書!B10': 1 } },
  { src: 'tests/fixtures/keihi_seisan.xlsx', data: { 明細: Array(6).fill({}) }, errors: true },
  { src: 'tests/fixtures/keihi_seisan.xlsx', data: { Y15: 1 }, errors: true },
  { src: 'tests/fixtures/keihi_seisan.xlsx', data: { 存在しない欄: 1 } },
  { src: 'tests/fixtures/keihi_seisan.xlsx', data: { 支払方法: '小切手' }, errors: true },
  { src: 'tests/fixtures/plain_colon.xlsx', data: { 氏名: '鈴木', 'Sheet1!D10': 'extra', 生年月日: '2000-01-02' },
    expect: { 'Sheet1!B3': '鈴木', 'Sheet1!D10': 'extra', 'Sheet1!B5': '2000年1月2日' } },
  // tests/test_hard.py
  { src: 'tests/fixtures_hard/rirekisho.xlsx', data: {
    記入日: '2026-11-19', '氏名(ふりがな)': 'やまだ たろう', 氏名: '山田 太郎', 生年月日: '1990-04-02', 性別: '男', '現住所(〒)': '123-4567',
    現住所: '東京都千代田区1-1', '学歴・職歴': [{ 年: 2009, 月: 4, '学歴・職歴': '〇〇大学 入学' }] },
    expect: { '履歴書!Q3': '令和8年11月19日現在', '履歴書!F7': '平成2年4月2日生（満36歳）', '履歴書!W7': '（男） ・ 女', '履歴書!F9': '〒123－4567', '履歴書!G14': '〇〇大学 入学' } },
  { src: 'tests/fixtures_hard/juminhyo.xlsx', data: {
    '請求者/生年月日': '1985-06-15', '必要な方/生年月日': '2020-01-10', 請求者との関係: { その他: '祖父' }, 使用目的: 'その他（パスポート申請）', 必要通数: 2 },
    expect: { '請求書!W11': '昭和60年6月15日', '請求書!I15': '令和2年1月10日', '請求書!I16': '□本人　□同一世帯の方　■その他（祖父）', '請求書!I19': '□就職　□年金　□保険　■その他（パスポート申請）' } },
  { src: 'tests/fixtures_hard/shinsei_kikan.xlsx', data: {
    申請日: '2026-11-19', 電話番号: '03-1234-5678', 出張期間: '2026-12-01～2026-12-03', 出発時刻: '9:30', 帰着時刻: [18, '05'],
    口座番号: '1234567', 支店: '渋谷', 口座種別: '普通', 概算金額: 52000 },
    expect: { '申請書!H7': '（03）1234－5678', '申請書!H8': '令和8年12月1日　～　令和8年12月3日（3日間）', '申請書!H9': '9時30分', '申請書!X9': '18時05分',
      '申請書!H16': 1, '申請書!I16': 2, '申請書!J16': 3, '申請書!K16': 4, '申請書!L16': 5, '申請書!M16': 6, '申請書!N16': 7, '申請書!T14': '銀行・信用金庫渋谷支店' } },
  { src: 'tests/fixtures_hard/multisheet.xlsx', data: { 備考: '駐車場利用', H6: '総務部', '本紙!来訪目的': '工事' },
    expect: { '本紙!H9': '駐車場利用', '本紙!H6': '総務部' } },
  { src: 'tests/fixtures_hard/kintou.xlsx', data: {
    宛名: '株式会社〇〇', 氏名: '佐藤 花子', 年齢: 28, 勤続年数: '5年3ヶ月', 電話番号: '090-1234-5678', 性別: '女性', 郵便番号: '1000001' },
    expect: { '申込書!H8': 5, '申込書!K8': 3, '申込書!H9': '090', '申込書!M9': 1234, '申込書!R9': 5678, '申込書!H10': '□', '申込書!L10': '■' } },
  { src: 'tests/fixtures_hard/kintou.xlsx', data: { 氏名: '佐藤 花子' }, expect: { '申込書!H6': '佐藤 花子' } },
  { src: 'tests/fixtures_hard/underline.xlsx', data: { 住所: '東京都', 氏名: '山田' } },
  // tests/test_hard.py DogfoodRegressionTests
  { src: 'tests/fixtures_hard/shinsei_kikan.xlsx', data: { 申請日: 'YYYY-MM-DD', 所属: '', 氏名: null, 学歴: [] }, noWrites: true },
  { src: 'tests/fixtures_hard/shinsei_kikan.xlsx', data: { 申請日: '来週' }, errors: true },
  { src: 'tests/fixtures_hard/shinsei_kikan.xlsx', data: { 出張期間: '2026-12-01/2026-12-03', 出発時刻: '09:30', 概算金額: '52000' },
    expect: { '申請書!H8': '令和8年12月1日　～　令和8年12月3日（3日間）', '申請書!H9': '9時30分', '申請書!H11': 52000 } },
  { src: 'tests/fixtures_hard/shinsei_kikan.xlsx', data: { 出張期間: '2026-12-01' }, warnings: true },
  { src: 'tests/fixtures_hard/rirekisho.xlsx', data: { 現住所: '東京都千代田区1-1', 性別: '女', '学歴・職歴': [{ 年: '2009', 月: '4', '学歴・職歴': '入学' }] },
    expect: { '履歴書!B14': 2009 }, readback: ['| 現住所 | テキスト | F10:Y10 | [履歴書] 現在値: 東京都千代田区1-1 |', '1行記入済み', '現在値: 男 ・ （女）'] },
  // real-form layouts and Excel form-control check boxes
  { src: 'tests/fixtures_hard/frame_todoke.xlsx', data: { 日付: '2026-11-19', 法人番号: '1234567890123', 担当者名: '山田', '所在地（住所）': '広島市' },
    expect: { '届出書!D15': 8, '届出書!G15': 11, '届出書!J15': 19, '届出書!O12': '山田' } },
  { src: 'tests/fixtures_hard/frame_todoke.xlsx', data: { 日付: '2018-01-01' }, errors: true },
  { src: 'tests/fixtures_hard/frame_todoke.xlsx', data: { 送付先: '730-0011 広島市中区' }, expect: { '届出書!F17': '〒730-0011 広島市中区' } },
  { src: 'tests/fixtures_hard/frame_todoke.xlsx', data: { 変更年月日: '2026-12-01' }, expect: { '届出書!R21': 8, '届出書!V21': 12, '届出書!Y21': 1 } },
  { src: 'tests/fixtures_hard/checkbox_controls.xlsx', data: { '名称変更/社名変更': true, '名称変更/法人成り': false, '所在地変更/その他': 'はい', 氏名: '山田' },
    expect: { '変更届!F8': '山田' } },
  { src: 'tests/fixtures_hard/checkbox_controls.xlsx', data: { '名称変更/社名変更': 'たぶん' }, errors: true },
  { src: 'tests/fixtures_hard/henko_todoke.xlsx', data: { '所在地（住所）/変更後': '広島市南区1-2-3', '所在地（住所）(〒)/変更後': '732-0001', '電話番号/変更後': '082-111-2222' },
    expect: { '変更届!R10': '広島市南区1-2-3', '変更届!S9': '732-0001', '変更届!S11': '082', '変更届!V11': 111, '変更届!Y11': 2222 } },
  // tests/test_safety.py (fixtures_review: files saved by real Excel)
  { src: 'tests/fixtures_review/base.xlsx', data: { A11: 'x', A12: 'y', A14: 'z' }, expect: { 'Sheet1!A11': 'x', 'Sheet1!A12': 'y', 'Sheet1!A14': 'z' } },
  { src: 'tests/fixtures_review/empty.xlsx', data: { A1: 'a', B3: 'b', C2: 'c' }, expect: { 'Sheet1!A1': 'a', 'Sheet1!B3': 'b', 'Sheet1!C2': 'c' } },
  { src: 'tests/fixtures_review/base.xlsx', data: { A0: 'x' }, errors: true },
  { src: 'tests/fixtures_review/base.xlsx', data: { A1048577: 'x' }, errors: true },
  { src: 'tests/fixtures_review/base.xlsx', data: { XFE1: 'x' }, errors: true },
  { src: 'tests/fixtures_review/base.xlsx', data: { B1: 'x'.repeat(32768) }, errors: true },
  { src: 'tests/fixtures_review/base.xlsx', data: { B1: 'a￾b￿c\ud800d' }, expect: { 'Sheet1!B1': 'abcd' } },
  { src: 'tests/fixtures_review/table.xlsx', data: { A1: '氏名' }, errors: true },
  { src: 'tests/fixtures_review/table.xlsx', data: { A3: 'x', A4: 'y' }, expect: { 'Sheet1!A3': 'x', 'Sheet1!A4': 'y' } },
  { src: 'tests/fixtures_review/norowrefs.xlsx', data: { A3: 'x' }, errors: true },
  { src: 'tests/fixtures_review/colstyle.xlsx', data: { C5: 'x', E7: 'y' }, expect: { 'Sheet1!C5': 'x', 'Sheet1!E7': 'y' } },
  { src: 'tests/fixtures_review/base.xlsx', data: { B1: 'ID_x0031_23' }, expect: { 'Sheet1!B1': 'ID_x0031_23' } },
  { src: 'tests/fixtures_review/formulas.xlsx', data: { D2: 999 }, errors: true },
  { src: 'tests/fixtures_review/spill.xlsx', data: { A2: 99 }, errors: true },
  { src: 'tests/fixtures/kyuka_todoke.xlsx', data: { I9: 'x' }, errors: true },
  { src: 'tests/fixtures_review/base.xlsx', data: { TEL1: '03-1234-5678' } },
  { src: 'tests/fixtures_review/prefixed.xlsx', data: { B2: 'x', C9: 1 }, expect: { 'Sheet1!B2': 'x', 'Sheet1!C9': 1 } },
  { src: 'tests/fixtures_hard/ringi.xlsx', data: { 部長: '済', 件名: 'PC購入', 金額: '1,200', 添付書類: ['見積書', 'カタログ'] } },
];
// plus one auto-generated case per fixture from its template
for (const file of fixtures) {
  const wb = await M.load(readFileSync(file));
  const fields = M.detect(wb);
  const data = {};
  for (const f of fields) {
    if (f.kind === 'table') data[f.id] = [Object.fromEntries(f.columns.slice(1).map((c, i) => [c.id, i ? 'x' : 100]))];
    else if (f.kind === 'choice' || f.kind === 'circle' || f.kind === 'choice_cells') data[f.id] = f.options[f.options.length - 1];
    else if (f.kind === 'date' || f.kind === 'inline_date') data[f.id] = '2026-04-01';
    else if (f.kind === 'digits') data[f.id] = '1234567'.slice(0, f.parts.length);
    else if (f.kind === 'split') data[f.id] = f.parts.map((_, i) => String(i + 1));
    else if (f.kind === 'pattern') data[f.id] = '12';
    else data[f.id] = '記入テスト';
  }
  fillCases.push({ src: path.relative(ROOT, file), data, auto: true });
}

async function runFillCases(cases, tag) {
  const jsJobs = [], pyJobs = [];
  for (const [i, c] of cases.entries()) {
    const src = path.isAbsolute(c.src) ? c.src : FX(c.src);
    const jsOut = path.join(TMP, `${tag}${i}_js.xlsx`), pyOut = path.join(TMP, `${tag}${i}_py.xlsx`);
    for (const p of [jsOut, pyOut]) if (existsSync(p)) rmSync(p);
    jsJobs.push({ c, src, jsOut, pyOut, res: await jsFill(src, jsOut, c.data, c.mark || '■') });
    pyJobs.push({ src, dst: pyOut, data: c.data, mark: c.mark || '■' });
  }
  const pyRes = helper('fill', pyJobs);
  const cmp = [];
  jsJobs.forEach((j, i) => {
    const name = `${path.basename(j.src)}#${i}${j.c.auto ? ' (auto)' : ''}`;
    const p = pyRes[i];
    if (p.exception || j.res.exception) {
      check(!!p.exception === !!j.res.exception, `${name}: exception mismatch js=${j.res.exception} py=${p.exception}`);
      return;
    }
    const a = normReport(j.res.report), b = normReport(p.report);
    if (!check(same(a, b), `${name}: report differs — ${diffLine(a, b)}`)) console.log('    js errors:', j.res.report.errors, '\n    py errors:', p.report.errors);
    if (j.c.errors !== undefined) check(!!b.errors.length === j.c.errors, `${name}: expected errors=${j.c.errors}`);
    if (j.c.errors) check(!existsSync(j.jsOut) && !existsSync(j.pyOut), `${name}: no file may be written when refusing`);
    if (j.c.warnings) check(b.warnings.length > 0, `${name}: expected warnings`);
    if (j.c.noWrites) check(b.ok && same(b.written, []), `${name}: expected nothing written`);
    if (!b.errors.length && existsSync(j.jsOut) && existsSync(j.pyOut)) cmp.push({ name, a: j.pyOut, b: j.jsOut, jsOut: j.jsOut, c: j.c });
  });
  if (cmp.length) {
    const diffs = helper('compare', cmp.map((x) => ({ a: x.a, b: x.b })));
    cmp.forEach((x, i) => check(Array.isArray(diffs[i]) && diffs[i].length === 0, `${x.name}: output parts differ from Python's: ${JSON.stringify(diffs[i])}`));
  }
  return cmp;
}
const done = await runFillCases(fillCases, 'case');
// literal Python-loader read-back of the JS outputs for the asserted cells
for (const x of done) {
  if (!x.c.expect) continue;
  const cells = Object.keys(x.c.expect);
  const script = 'import sys, json; sys.path.insert(0, "src"); from masume import load; from masume.ooxml import parse_ref\n' +
    'wb = load(sys.argv[1]); cells = json.loads(sys.argv[2])\n' +
    'sys.stdout.buffer.write(json.dumps([wb.sheet(c.rsplit("!",1)[0]).value(*parse_ref(c.rsplit("!",1)[1])) for c in cells], ensure_ascii=False).encode("utf-8"))';
  const got = JSON.parse(py(['-c', script, x.jsOut, JSON.stringify(cells)]));
  const want = cells.map((c) => x.c.expect[c]);
  check(same(got, want), `${x.name}: Python loader read ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
}
// readback: `python -m masume md --fields` vs JS workbookToMarkdown(wb, layout, savedFields)
for (const x of done) {
  if (!x.c.readback) continue;
  const fieldsJson = path.join(TMP, 'fields.json');
  writeFileSync(fieldsJson, py(['-m', 'masume', 'inspect', '--json', FX(x.c.src)]));
  for (const layout of [false, true]) {
    const pyMd = py(['-m', 'masume', 'md', '--fields', fieldsJson, ...(layout ? [] : ['--no-layout']), x.jsOut]);
    const saved = JSON.parse(readFileSync(fieldsJson, 'utf8')).map((d) => M.Field.fromDict(d));
    const jsMd = M.workbookToMarkdown(await M.load(readFileSync(x.jsOut)), layout, saved) + '\n';
    check(jsMd === pyMd, `${x.name}: md --fields (layout=${layout}) differs — ${diffLine(jsMd.split('\n'), pyMd.split('\n'))}`);
    for (const needle of x.c.readback) check(jsMd.includes(needle), `${x.name}: readback lacks ${needle}`);
  }
}
// values JSON cannot carry: checked on the JS side only
{
  const base = await M.load(readFileSync(FX('tests/fixtures_review/base.xlsx')));
  for (const v of [NaN, Infinity, -Infinity]) {
    const { report, bytes } = await M.fill(base, { B1: v });
    check(report.errors.some((e) => e.includes('cannot store')) && bytes === null, `B1=${v} must be refused`);
  }
}
console.log(`  ${fillCases.length} fill cases, ${done.length} outputs byte-compared with Python's`);

// ============================================================================
console.log(`4. random workbooks (${FUZZ}): detection + markdown + fill parity`);
function rng(seed) {
  let a = seed >>> 0;
  const f = () => { a = (a + 0x6D2B79F5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
  f.int = (n) => Math.floor(f() * n);
  f.pick = (arr) => arr[Math.floor(f() * arr.length)];
  f.chance = (p) => f() < p;
  return f;
}
const LABELS = ['氏名', '所　属', '氏　　名', '氏名：', 'メール：', '生年月日', '生年月日：', '日付', '期日', '開始日', '終了日', '申請日',
  '提出日', '内容', '金額', '数量', '備考', '（備考）', '【申請者】', '※注意事項', '合計', '金額*', '所属部署名', '電話番号', '住所',
  '理由', '社員番号', '𠮷田', 'a|b', '改行\nあり', ' 前後空白 ', '品名', '区分', 'ふりがな', 'フリガナ', '郵便番号', '年齢', '勤続年数',
  '口座番号', '性別', '■ 申込者情報', '第1章 概要', '請求者', '必要な方', '宛名', '出発時刻', '支店', '連絡先', 'E-mail', '決裁', '部長', '課長'];
const SPECIAL = ['□有　□無', '■有 □無', '☐ 男性 ☐ 女性', '□本人　□その他（　　）', '□現金　□振込（□普通 □当座）', '✓済', '□', '■', '☑',
  '令和　　年　　月　　日', '令和　年　月　日', '提出日：令和　年　月　日', '西暦　　年　　月　　日', '　　年　　月　　日', '平成２年１月３日',
  '令和5年4月1日', '年月日', '昭和・平成　　年　　月　　日生（満　　歳）', '明・大・昭・平・令　　年　　月　　日', '令和　年　月　日現在',
  '令和　年　月　日　～　令和　年　月　日（　　日間）', '備考：＿＿＿＿＿', '氏名 ____', '電話番号＿＿＿＿＿＿＿（自宅）', '__', '100', '1,000円',
  '¥500', '１２', '12.5', 'これはとても長いラベルテキストで四十文字を超えるように書かれた説明文ですよね本当に長いのです', '男 ・ 女', '有・無',
  '可/否', '〒　　　－', '（　　　）　　　－', '　　時　　分', '銀行・信用金庫　　　支店', '様', '御中', '殿', '円', '歳', '年', 'ヶ月', '－',
  '日間', '〒', '印', '・', 'No', 'No.', '番号', '#', '№', 'ｎｏ', '日', '小計'];
const TABLE_HEADS = ['No', '日付', '内容', '金額', '数量', '品名', '備考', '区分', '年', '月', '学歴・職歴'];
const NUMS = ['1', '2', '3.0', '1E3', '2.5', '-4', '0', ' 7 ', '１２', '1_000', 'abc'];
const SHEETNAMES = ['Sheet1', '申請書', 'データ', '記入例', '本紙', 'Form'];
const xesc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function genSheet(r, shared) {
  const R = 8 + r.int(30), C = 8 + r.int(26);
  const occ = new Set();
  const cells = new Map(); // key -> {s, xml(p)}
  const merges = [];
  const free = (r1, c1, r2, c2) => { for (let y = r1; y <= r2; y++) for (let x = c1; x <= c2; x++) if (occ.has(y * 1000 + x) || x > C || y > R) return false; return true; };
  const take = (r1, c1, r2, c2) => { for (let y = r1; y <= r2; y++) for (let x = c1; x <= c2; x++) occ.add(y * 1000 + x); };
  const strCell = (text) => {
    const k = r.int(10);
    if (k < 6) {
      let idx = shared.indexOf(text);
      if (idx < 0 || r.chance(0.2)) {
        let si;
        if (r.chance(0.2) && text.length > 1) si = `<r><t xml:space="preserve">${xesc(text.slice(0, 1))}</t></r><r><rPr><b/></rPr><t xml:space="preserve">${xesc(text.slice(1))}</t></r>`;
        else si = `<t xml:space="preserve">${xesc(text)}</t>`;
        if (r.chance(0.25)) si += '<rPh sb="0" eb="1"><t>フリ</t></rPh><phoneticPr fontId="0"/>';
        shared.push(text); idx = shared.length - 1; shared.xml.push(`<si>${si}</si>`);
      }
      return { t: 's', inner: (p) => `<${p}v>${idx}</${p}v>` };
    }
    if (k < 9) return { t: 'inlineStr', inner: (p) => `<${p}is><${p}t xml:space="preserve">${xesc(text)}</${p}t></${p}is>` };
    return { t: 'str', inner: (p) => `<${p}v>${xesc(text)}</${p}v>` };
  };
  const valueCell = (kind) => {
    if (kind === 'label') return strCell(r.pick(LABELS));
    if (kind === 'special') return strCell(r.pick(SPECIAL));
    if (kind === 'num') { const v = r.pick(NUMS); return { t: null, inner: (p) => `<${p}v>${v}</${p}v>` }; }
    if (kind === 'bool') { const v = r.int(2); return { t: 'b', inner: (p) => `<${p}v>${v}</${p}v>` }; }
    if (kind === 'formula') {
      const k = r.int(4);
      if (k === 0) return { t: null, inner: (p) => `<${p}f>SUM(A1:A2)</${p}f><${p}v>3</${p}v>` };
      if (k === 1) return { t: null, inner: (p) => `<${p}f>SUM(A1:A2)</${p}f>` };
      if (k === 2) return { t: 'str', inner: (p) => `<${p}f>A1&amp;"x"</${p}f><${p}v>abc</${p}v>` };
      return { t: 'e', inner: (p) => `<${p}f>1/0</${p}f><${p}v>#DIV/0!</${p}v>` };
    }
    return null;
  };
  const putCell = (y, x, s, v) => cells.set(y * 1000 + x, { y, x, s, v });
  const region = (r1, c1, r2, c2, kind, style) => {
    take(r1, c1, r2, c2);
    if (r1 !== r2 || c1 !== c2 || r.chance(0.03)) merges.push(`${M.makeRef(r1, c1)}:${M.makeRef(r2, c2)}`);
    const v = valueCell(kind);
    const mode = r.int(3); // style on: all cells / top-left / outline
    for (let y = r1; y <= r2; y++) for (let x = c1; x <= c2; x++) {
      const tl = y === r1 && x === c1;
      const on = mode === 0 || tl || (mode === 2 && (y === r1 || y === r2 || x === c1 || x === c2));
      if (tl) putCell(y, x, style, v);
      else if (on && style) putCell(y, x, style, null);
    }
  };
  const boxStyle = () => r.pick([1, 1, 1, 2, 4, 6, 0]);
  const labStyle = () => r.pick([0, 0, 1, 5, 7, 6]);
  for (let y = 1; y <= R; y++) {
    // motifs
    if (r.chance(0.08) && y + 4 <= R) { // table: header row + body rows with one empty row
      const n = 2 + r.int(3);
      let x = 1 + r.int(3);
      const spans = [];
      for (let i = 0; i < n; i++) { const w = 1 + r.int(5); spans.push([x, x + w - 1]); x += w; }
      if (x - 1 <= C && free(y, 1, y + 4, x)) {
        if (r.chance(0.5) && y > 1 && free(y - 1, spans[0][0], y - 1, spans[n - 1][1])) region(y - 1, spans[0][0], y - 1, spans[n - 1][1], 'label', 0);
        spans.forEach(([a, b], i) => { take(y, a, y, b); merges.push(`${M.makeRef(y, a)}:${M.makeRef(y, b)}`); putCell(y, a, 1, strCell(i === 0 && r.chance(0.4) ? 'No' : r.pick(TABLE_HEADS))); });
        const rows = 2 + r.int(3);
        for (let k = 1; k <= rows; k++) spans.forEach(([a, b], i) => region(y + k, a, y + k, b, k === 1 && r.chance(0.3) ? (i === 0 ? 'num' : 'label') : (k === rows && r.chance(0.1) ? 'formula' : 'empty'), 1));
        y += rows;
        continue;
      }
    }
    let x = 1;
    while (x <= C) {
      if (r.chance(0.25)) { x += 1 + r.int(2); continue; }
      const motif = r.int(12);
      if (motif === 0 && free(y, x, y, x + 8)) { // label + split boxes with units
        region(y, x, y, x + 1, 'label', labStyle());
        region(y, x + 2, y, x + 3, 'empty', 1); region(y, x + 4, y, x + 4, 'special', 0);
        region(y, x + 5, y, x + 6, 'empty', 1); region(y, x + 7, y, x + 7, 'special', 0);
        x += 9; continue;
      }
      if (motif === 1 && free(y, x, y, x + 9)) { // label + single-cell digit / underline run
        region(y, x, y, x + 1, 'label', labStyle());
        const st = r.pick([1, 2, 4]);
        const n = 2 + r.int(6);
        for (let k = 0; k < n; k++) region(y, x + 2 + k, y, x + 2 + k, 'empty', st);
        x += 3 + n; continue;
      }
      if (motif === 2 && free(y, x, y, x + 7)) { // label + □ cells
        region(y, x, y, x + 1, 'label', labStyle());
        for (let k = 0; k < 2 + r.int(2); k++) {
          if (!free(y, x + 2 + 2 * k, y, x + 3 + 2 * k)) break;
          take(y, x + 2 + 2 * k, y, x + 3 + 2 * k);
          putCell(y, x + 2 + 2 * k, 0, strCell(r.pick(['□', '□', '■', '☐'])));
          putCell(y, x + 3 + 2 * k, 0, strCell(r.pick(['男性', '女性', '有', '無', 'その他', '□'])));
        }
        x += 9; continue;
      }
      if (motif === 3 && y + 3 <= R && free(y, x, y + 3, x + 6)) { // tall narrow group label
        region(y, x, y + 3, x + r.int(3), 'label', 1);
        x += 4; continue;
      }
      const w = 1 + r.int(9), hgt = r.chance(0.82) ? 1 : 2 + r.int(3);
      if (!free(y, x, y + hgt - 1, x + w - 1)) { x += 1; continue; }
      const kind = r.pick(['empty', 'empty', 'label', 'label', 'label', 'special', 'special', 'num', 'bool', 'formula']);
      region(y, x, y + hgt - 1, x + w - 1, kind, kind === 'empty' ? boxStyle() : r.chance(0.5) ? boxStyle() : labStyle());
      x += w;
    }
  }
  return { cells, merges, R };
}
function sheetXml(r, sh, prefix) {
  const p = prefix;
  const rows = new Map();
  for (const c of sh.cells.values()) { if (!rows.has(c.y)) rows.set(c.y, []); rows.get(c.y).push(c); }
  let body = '';
  for (const y of [...rows.keys()].sort((a, b) => a - b)) {
    const noR = r.chance(0.02);
    const cs = rows.get(y).sort((a, b) => a.x - b.x);
    body += `<${p}row r="${y}" spans="1:${cs[cs.length - 1].x}"${r.chance(0.2) ? ' ht="12" customHeight="1"' : ''}>`;
    for (const c of cs) {
      const rAttr = noR ? '' : ` r="${M.makeRef(c.y, c.x)}"`;
      const sAttr = c.s ? ` s="${c.s}"` : '';
      if (!c.v) body += `<${p}c${rAttr}${sAttr}/>`;
      else body += `<${p}c${rAttr}${sAttr}${c.v.t ? ` t="${c.v.t}"` : ''}>${c.v.inner(p)}</${p}c>`;
    }
    body += `</${p}row>`;
  }
  const ns = p ? `xmlns:${p.slice(0, -1)}="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` : 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  const sd = body ? `<${p}sheetData>${body}</${p}sheetData>` : `<${p}sheetData/>`;
  const mc = sh.merges.length ? `<${p}mergeCells count="${sh.merges.length}">${sh.merges.map((m) => `<${p}mergeCell ref="${m}"/>`).join('')}</${p}mergeCells>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<${p}worksheet ${ns} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><${p}dimension ref="A1"/>${sd}${mc}<${p}pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></${p}worksheet>`;
}
async function genWorkbook(seed) {
  const r = rng(seed);
  const shared = []; shared.xml = [];
  const nSheets = 1 + (r.chance(0.3) ? 1 + r.int(2) : 0);
  const names = [];
  const files = [];
  const sheetsXml = [];
  const rels = [];
  for (let i = 1; i <= nSheets; i++) {
    let name = r.pick(SHEETNAMES);
    while (names.includes(name)) name += 'x';
    names.push(name);
    const sh = genSheet(r, shared);
    files.push({ name: `xl/worksheets/sheet${i}.xml`, data: sheetXml(r, sh, r.chance(0.25) ? 'x:' : '') });
    const hidden = i > 1 && r.chance(0.3) ? ' state="hidden"' : '';
    sheetsXml.push(`<sheet name="${xesc(name)}" sheetId="${i}"${hidden} r:id="rId${i}"/>`);
    rels.push(`<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${r.chance(0.2) ? '/xl/' : ''}worksheets/sheet${i}.xml"/>`);
  }
  rels.push('<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>');
  rels.push('<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>');
  const calc = r.pick(['', '<calcPr calcId="191029"/>', '<calcPr fullCalcOnLoad="0" calcId="1"/>']);
  const defined = r.chance(0.3) ? '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$C$3</definedName></definedNames>' : '';
  files.push({ name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>' });
  files.push({ name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' });
  files.push({ name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetsXml.join('')}</sheets>${defined}${calc}</workbook>` });
  files.push({ name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>` });
  files.push({ name: 'xl/styles.xml', data: '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="3"><font><sz val="11"/></font><font><b/><sz val="11"/></font><font><sz val="16"/></font></fonts><fills count="1"><fill/></fills>' +
    '<borders count="5"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border><border><bottom style="hair"/></border><border><left style="none"/></border><border><left style="thin"/><right style="thin"/></border></borders>' +
    '<cellXfs count="8"><xf borderId="0" fontId="0"/><xf borderId="1" fontId="0"/><xf borderId="2" fontId="0"/><xf borderId="3" fontId="0"/><xf borderId="4" fontId="0"/><xf borderId="0" fontId="1"/><xf borderId="1" fontId="1"/><xf borderId="0" fontId="2"/></cellXfs></styleSheet>' });
  files.push({ name: 'xl/sharedStrings.xml', data: `<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${shared.length}" uniqueCount="${shared.length}">${shared.xml.join('')}</sst>` });
  files.push({ name: 'xl/media/image1.png', data: new Uint8Array([0x89, 0x50, 0x4E, 0x47, 1, 2, 3, 4]) });
  return { bytes: await M.zipFiles(files, r.chance(0.5)), names };
}
function genData(r, fields, names) {
  const data = {};
  const dates = ['2026-11-19', '2019-05-01', '2019-04-30', '1989/1/8', '1989-01-07', '2026.1.2', '２０２０年１月２日', ' 2020-02-29 ', '2021-02-29', '未定', 20261119, null];
  for (const f of fields) {
    if (r.chance(0.2)) continue;
    const key = r.chance(0.15) ? f.label : f.id;
    let v;
    switch (f.kind) {
      case 'table': {
        const n = r.chance(0.05) ? f.rows.length + 1 : r.int(f.rows.length + 1);
        v = Array.from({ length: n }, () => (r.chance(0.3)
          ? f.columns.map(() => r.pick(['交通費', 2300, '', null, '11/1']))
          : Object.fromEntries(f.columns.filter(() => r.chance(0.6)).map((c) => [r.chance(0.2) ? c.label : c.id, r.pick(['交通費', 2300, '', null, '11/1', 3.5])]))));
        break;
      }
      case 'date': case 'inline_date':
        v = r.chance(0.2) ? r.pick([['2026-12-01', '2026-12-03'], '2026-12-01～2026-12-03', { from: '2026-12-01', to: '2026-12-05' }, '2026-12-01 to 2026-12-02'])
          : r.pick(dates);
        break;
      case 'choice': {
        const o = f.options;
        v = r.pick([o[0], o[o.length - 1], [o[0], o[o.length - 1]], `${o[o.length - 1]}（メモ）`, { [o[0]]: true }, { [o[o.length - 1]]: 'メモ' }, r.chance(0.2) ? 'ないもの' : o[0], null]);
        break;
      }
      case 'circle': v = r.pick([...f.options, [f.options[0]], r.chance(0.2) ? 'ないもの' : f.options[0]]); break;
      case 'choice_cells': v = r.pick([...f.options, f.options.slice(0, 2), r.chance(0.2) ? 'ないもの' : f.options[0]]); break;
      case 'pattern': v = r.pick(['123-4567', '03-1234-5678', '9:30', [1, 2], '渋谷', 'a b c d e', 1000001, '']); break;
      case 'split': v = r.pick(['5年3ヶ月', [1, 2], ['1', '2', '3'], { 年: 5, ヶ月: '3' }, '1000001', '12', '090-1234-5678']); break;
      case 'digits': v = r.pick(['1234567', '12-34', ['1', '2'], '１２３', 'ABC', '0012']); break;
      case 'inline_text': v = r.pick(['なし', 'テスト記入', 123, '']); break;
      default: v = r.pick(['山田 太郎', 42, '0123', 3.5, true, null, '', 'a<b&c>"\'', '改行\nあり', '  前後  ', -7, '𠮷', '2026-01-02', 'x\u0001y']);
    }
    data[key] = v;
  }
  if (r.chance(0.4)) data[r.pick(['B2', '$D$4', 'C3', 'A1', 'AA30'])] = 'セル指定';
  if (r.chance(0.2)) data[`${r.pick(names)}!B3`] = 7;
  if (r.chance(0.15)) data['なにか'] = 1;
  if (r.chance(0.1)) data['Nosheet!A1'] = 1;
  return data;
}
function validData(r, fields) {
  const data = {};
  for (const f of fields) {
    if (r.chance(0.15)) continue;
    let v;
    switch (f.kind) {
      case 'table': v = Array.from({ length: 1 + r.int(f.rows.length) }, () => Object.fromEntries(f.columns.slice(1).map((c) => [c.id, r.pick(['品目', 120, '11/1', 2.5])]))); break;
      case 'date': case 'inline_date': v = r.pick(['2026-11-19', '2019-05-01', '1989-01-08', '2000/2/29', '２０２０年１月２日']); break;
      case 'choice': v = r.chance(0.3) ? f.options.slice(0, 2) : r.pick(f.options); break;
      case 'circle': v = r.pick(f.options); break;
      case 'choice_cells': v = r.chance(0.3) ? f.options : r.pick(f.options); break;
      case 'pattern': v = r.pick(['12', '渋谷']); break;
      case 'split': v = f.parts.map((_, k) => r.pick([String(k + 1), '0123', k * 10])); break;
      case 'digits': v = '9876543210'.slice(0, 1 + r.int(f.parts.length)); break;
      case 'inline_text': v = r.pick(['なし', 'テスト']); break;
      default: v = r.pick(['山田 太郎', 42, '0123', 3.5, true, '', 'a<b&c>', '改行\nあり', '  前後  ']);
    }
    data[f.id] = v;
  }
  return data;
}
{
  const dir = path.join(TMP, 'fuzz');
  mkdirSync(dir);
  const files = [], meta = [];
  for (let i = 0; i < FUZZ; i++) {
    const seed = 1000 + i;
    const { bytes, names } = await genWorkbook(seed);
    const f = path.join(dir, `fz${seed}.xlsx`);
    writeFileSync(f, bytes);
    files.push(f);
    meta.push({ seed, names });
  }
  const pyAll = files.length ? helper('inspect', files) : [];
  let detOk = 0;
  const kinds = {};
  const cases = [];
  for (let i = 0; i < files.length; i++) {
    const p = pyAll[i];
    let js;
    try { js = await jsInspect(files[i]); } catch (e) { js = { exception: `${e.constructor.name}: ${e.message}` }; }
    if (p.exception || js.exception) { check(!!p.exception === !!js.exception, `fz${meta[i].seed}: load exception mismatch js=${js.exception} py=${p.exception}`); continue; }
    for (const f of p.fields) kinds[f.kind] = (kinds[f.kind] || 0) + 1;
    const a = check(same(js.fields, p.fields), `fz${meta[i].seed}: fields differ — ${diffLine(js.fields, p.fields)}`);
    const b = check(js.md === p.md, `fz${meta[i].seed}: markdown differs — ${diffLine(js.md.split('\n'), p.md.split('\n'))}`);
    const c = check(same(js.template, p.template), `fz${meta[i].seed}: template differs`);
    if (a && b && c) {
      detOk++;
      const wb = await M.load(readFileSync(files[i]));
      const fields = M.detect(wb);
      cases.push({ src: files[i], data: genData(rng(meta[i].seed * 7 + 1), fields, meta[i].names), mark: rng(meta[i].seed).chance(0.3) ? '☑' : '■' });
      cases.push({ src: files[i], data: validData(rng(meta[i].seed * 13 + 5), fields) });
    }
  }
  console.log(`  detection identical for ${detOk}/${files.length} workbooks; field kinds seen: ${JSON.stringify(kinds)}`);
  const cmp = await runFillCases(cases, 'fz');
  console.log(`  ${cases.length} random fills compared, ${cmp.length} outputs byte-compared`);
}

if (!KEEP) rmSync(TMP, { recursive: true, force: true }); else console.log('kept temp files in', TMP);
console.log(failures ? `\n${failures} FAILED of ${checks} checks` : `\nall ${checks} checks passed`);
process.exit(failures ? 1 : 0);
