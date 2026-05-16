// 月次経営ダッシュボード メインアプリ
// データ: data/ground_truth_v1.json
// 構造: sheets[シート名].blocks[項目名][年度ラベル] = {"0": 4月値, ..., "11": 3月値}

const DATA_URL = "data/ground_truth_v1.json";
const MONTH_LABELS = ["4月","5月","6月","7月","8月","9月","10月","11月","12月","1月","2月","3月"];

// 法人 → 所属シート（xlsx 並び順）
const HOUJIN_SHEETS = {
  meiwakai: [
    "【自動】 統合)明和会+MS",
    "医）明和会　全体",
    "医）明和会　病院",
    "医）明和会　フェニックス",
    "医）明和会　保育園",
    "医）明和会　訪問看護",
    "医）明和会　在介",
    "医）明和会　GH",
    "医）明和会　居宅",
    "医）明和会　藍住居宅",
    "医）明和会　藍住クリニック",
    "医）明和会　デイケアあいそら",
    "医）明和会　藍住美容",
    "【自動】(統合)藍住クリニック＋美容＋デイケア",
  ],
  mediens: [
    "メディエンス　全体",
    "メディエンス　ハート徳島クリニック",
  ],
  ms: [
    "MS　全体",
    "MS　本社",
    "MS　GHふれあい",
    "MS　派遣",
    "MS 阿波っ子ショートステイ",
    "MS 阿波っ子デイ",
    "【自動】(統合)MS 阿波っ子ショート＋デイ",
    "MS 阿波っ子サ付高住",
    "MS 阿波っ子訪問ヘルパー",
    "【自動】(統合)MS 阿波っ子サ付＋訪問ヘルパー",
  ],
  shafuku: [
    "社福　全体",
    "社福　まごころ訪問ヘルパー",
    "社福　小規模特養あおぞら",
    "社福　はばたき認定こども園",
  ],
};

const HOUJIN_LABELS = {
  meiwakai: "（医）明和会",
  mediens: "（医）メディエンス",
  ms: "（有）たまきメディカルサポート",
  shafuku: "（社福）明和福祉会",
};

// PL項目表示順 — 正本: 知見集/18_大項目マッピング表.md
// 月次試算表.xlsx 病院シート準拠（経常利益を最上段に配置）
// 全法人（明和会・MS・社福・メディエンス）でこの順序を統一適用
const ITEM_ORDER = [
  // ① 経常利益（最上段・最重要KPI・xlsx R3）
  "経常利益",
  "経常損益",
  "明和会　経常利益",
  "MS　経常利益",
  "経常利益（明和会+MS）",

  // ② メイン収益（xlsx R41）
  "医業収益（全体）",
  "医業収益",
  "医療収益",
  "事業収益合計",
  "サービス活動収益",
  "売上高",
  "売上高（賃貸料収入）",
  "売上",
  "入居費収入＋売上高",
  "明和会 医業収益（全体）",
  "MS 売上高(全体)",
  "明和会医業収入＋MS 売上高",

  // ③ サブ収益（収益内訳・xlsx R79〜R155）
  "(入院診療収益)",
  "(外来診療収益)",
  "(自費診療)",
  "自費診療収入",
  "(健康診断等収益)",
  "健康診断等収益",
  "(入居介護収益)",
  "(デイ施設利用料)",
  "(自己負担収入)",
  "(保険収入）",
  "(訪看医療保険収入)",
  "(訪看介護保険収入)",
  "（賃貸料収入）",
  "（介護報酬収益）※R8/4期より入居介護収益／保険収入に科目変更",
  "（収入）",
  "（介護保険事業収益）",
  "（障害福祉サービス等収益）",

  // ④ 費用（xlsx R193→R268→R306→R344→R382→R420）
  // 順序：材料費 → 委託費 → 給与費 → 経費合計 → 設備関係費 → 減価償却費
  "売上原価",
  "売上総利益",
  "材料費合計",
  "材料費",
  "委託費合計",
  "委託費",
  "給与費合計",
  "給与費",
  "人件費",
  "明和会　給与費",
  "MS　給与費",
  "給与費（明和会＋MS）",
  "派遣料",
  "派遣料+給与",
  "経費合計",
  "販売費及び一般管理費計",
  "設備関係費合計",
  "減価償却費",

  // ⑤ 事業損益（xlsx R458）
  "事業損益",

  // ⑥ 事業外（xlsx R496→R534）
  "事業外収益",
  "事業外費用",

  // ⑦ 特別損益（xlsx R572→R610）
  "特別利益",
  "特別損失",

  // ⑧ 最終（xlsx R648→R686）
  "税引前当期純利益",
  "税引前当期純損益",
  "当期純利益",
  "当期純損益",
];
const ITEM_RANK = Object.fromEntries(ITEM_ORDER.map((k, i) => [k, i]));

function sortItems(items) {
  return items.slice().sort((a, b) => {
    const ra = ITEM_RANK[a] ?? 9999;
    const rb = ITEM_RANK[b] ?? 9999;
    if (ra !== rb) return ra - rb;
    return a.localeCompare(b, "ja");
  });
}

// 項目名の表示整形
// ① 後置修飾子「（全体）」削除
// ② 文字列全体が(xxx)で囲まれたサブ収益 → "xxx (個別)" に変換（階層を後置で示す）
// ③ 注釈(※)付き・補足括弧は触らない
function displayItemName(name) {
  if (!name) return name;
  // 注釈付きはそのまま（「（介護報酬収益）※...」など）
  if (name.includes("※")) return name;
  // 文字列全体が括弧で囲まれているサブ収益 → 中身 + " (個別)"
  const m = name.match(/^[（(](.+)[）)]$/);
  if (m) return m[1] + " (個別)";
  // 後置「（全体）」「(全体)」削除
  return name.replace(/[（(]全体[）)]\s*$/, "").trim();
}

// 短い施設名表示
// 【自動】(統合) / 【自動】 統合) → 【統合】 に統一
// 法人名プレフィックスは省略（タブ列の親文脈で自明）
function shortShisetsuName(s) {
  return s
    .replace(/^【自動】\s*[（(]?\s*統合\s*[）)]?\s*/, "【統合】")
    .replace(/^医）明和会　/, "")
    .replace(/^医）明和会/, "")
    .replace(/^MS　/, "")
    .replace(/^MS /, "")
    .replace(/^社福　/, "")
    .replace(/^メディエンス　/, "");
}

// 年度ラベル → 表示用 (例: "7.4-8.3" → "R7")
function yearLabelDisplay(label) {
  const m = label.match(/^(\d+)\.4-(\d+)\.3$/);
  if (!m) return label;
  const start = parseInt(m[1]);
  // 25-30はH(平成)、31以降は迷うが当院運用は31=H31、それ以降32+はR(令和)に
  // しかし試算表ではR1=元年=2019年なので：H25-H31 (2013-2019), R1-R7 (2019-2025)
  // 32.4-33.3 などはないので、>=32 → R(start-31)、<32 → そのまま和暦数値
  if (start >= 25 && start <= 31) return `H${start}`;
  return `R${start}`;
}

// 年度数値（古い→新しいで増える）。並び替え用
function yearOrderKey(label) {
  // "7.4-8.3" → 開始年(令和ベース→2019+7=2026), "31.4-2.3" → H31=2019, "2.4-3.3" → R2=2020
  const m = label.match(/^(\d+)\.4-(\d+)\.3$/);
  if (!m) return 0;
  const start = parseInt(m[1]);
  const end = parseInt(m[2]);
  // start=25..31 は平成 (西暦= start+1988)
  // start=1..30 で end<start は令和 (start=1なら2019,2なら2020,...,7なら2025)
  // start=31,end=2 のような跨ぎは H31 (2019)
  if (start >= 25 && start <= 31) return 1988 + start; // 平成
  return 2018 + start; // 令和
}

// 色生成（新しい年=濃い青、古い年=淡い青）
function yearColor(idx, total) {
  // idx=0 が最新（濃い）、idx=total-1 が最古（淡い）
  const t = total <= 1 ? 0 : idx / (total - 1);
  // HSL: hue 210 固定、lightness 25 → 88、saturation 70 → 35
  const l = 25 + t * 63;
  const s = 70 - t * 35;
  return `hsl(210, ${s}%, ${l}%)`;
}

// 値を取得（dict<str,number>）→ [m0,m1,...m11]
function monthArray(yearDict) {
  const arr = new Array(12).fill(null);
  for (const [k, v] of Object.entries(yearDict)) {
    const i = parseInt(k);
    if (i >= 0 && i < 12 && typeof v === "number") arr[i] = v;
  }
  return arr;
}

// 累積化
function cumulative(arr) {
  let acc = 0;
  let started = false;
  return arr.map(v => {
    if (v === null && !started) return null;
    started = true;
    acc += (v || 0);
    return acc;
  });
}

// 数値フォーマット
function fmt(n) {
  if (n === null || n === undefined) return "";
  return n.toLocaleString("ja-JP", { maximumFractionDigits: 0 });
}

// ---------- アプリ状態 ----------
const state = {
  data: null,
  houjin: null,
  shisetsu: null,
  koumoku: null,
  mode: "monthly", // monthly | cumulative
  chart: null,
};

// ---------- 画面切替 ----------
function showPage(id) {
  document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
  document.getElementById(id).classList.add("active");
}

// ---------- 法人選択画面 ----------
function bindHoujinButtons() {
  document.querySelectorAll(".houjin-card").forEach(btn => {
    btn.addEventListener("click", () => {
      const h = btn.dataset.houjin;
      enterHoujin(h);
    });
  });
  document.getElementById("logout").addEventListener("click", logout);
  document.getElementById("logout2").addEventListener("click", logout);
  document.getElementById("back-to-houjin").addEventListener("click", () => {
    showPage("page-houjin");
  });
}

function enterHoujin(houjin) {
  state.houjin = houjin;
  state.shisetsu = null;
  state.koumoku = null;
  document.getElementById("houjin-name").textContent = HOUJIN_LABELS[houjin];
  renderShisetsuTabs();
  // 最初の施設を選択
  const sheets = HOUJIN_SHEETS[houjin].filter(s => state.data.sheets[s]);
  if (sheets.length) selectShisetsu(sheets[0]);
  showPage("page-dashboard");
}

function renderShisetsuTabs() {
  const wrap = document.getElementById("shisetsu-tabs");
  wrap.innerHTML = "";
  const sheets = HOUJIN_SHEETS[state.houjin].filter(s => state.data.sheets[s]);
  sheets.forEach(s => {
    const btn = document.createElement("button");
    btn.className = "tab";
    btn.textContent = shortShisetsuName(s);
    btn.title = s;
    btn.dataset.sheet = s;
    btn.addEventListener("click", () => selectShisetsu(s));
    wrap.appendChild(btn);
  });
}

function selectShisetsu(sheetName) {
  state.shisetsu = sheetName;
  document.querySelectorAll("#shisetsu-tabs .tab").forEach(t => {
    t.classList.toggle("active", t.dataset.sheet === sheetName);
  });
  renderKoumokuTabs();
  const blocks = Object.keys(state.data.sheets[sheetName].blocks);
  const sorted = sortItems(blocks);
  if (sorted.length) selectKoumoku(sorted[0]);
}

function renderKoumokuTabs() {
  const wrap = document.getElementById("koumoku-tabs");
  wrap.innerHTML = "";
  const blocks = Object.keys(state.data.sheets[state.shisetsu].blocks);
  const sorted = sortItems(blocks);
  sorted.forEach(k => {
    const btn = document.createElement("button");
    btn.className = "tab";
    btn.textContent = displayItemName(k);
    btn.title = k;
    btn.dataset.item = k;
    btn.addEventListener("click", () => selectKoumoku(k));
    wrap.appendChild(btn);
  });
}

function selectKoumoku(item) {
  state.koumoku = item;
  document.querySelectorAll("#koumoku-tabs .tab").forEach(t => {
    t.classList.toggle("active", t.dataset.item === item);
  });
  renderChartAndTable();
}

// ---------- グラフ＆表 ----------
function getMatrix() {
  // 年度ラベル → 月別配列。新しい順に並べる
  const block = state.data.sheets[state.shisetsu].blocks[state.koumoku];
  const years = Object.keys(block).slice();
  years.sort((a, b) => yearOrderKey(b) - yearOrderKey(a)); // 新→古
  return years.map(y => ({
    yearLabel: y,
    display: yearLabelDisplay(y),
    values: monthArray(block[y]),
  }));
}

function renderChartAndTable() {
  const matrix = getMatrix();
  const useMode = state.mode;
  const dataSeries = matrix.map(row => useMode === "cumulative" ? cumulative(row.values) : row.values);

  // ---------- グラフ ----------
  const ctx = document.getElementById("main-chart").getContext("2d");
  if (state.chart) state.chart.destroy();

  const total = matrix.length;
  const datasets = matrix.map((row, idx) => ({
    label: row.display,
    data: dataSeries[idx],
    backgroundColor: yearColor(idx, total),
    borderColor: yearColor(idx, total),
    borderWidth: 1,
  }));

  state.chart = new Chart(ctx, {
    type: "bar",
    data: { labels: MONTH_LABELS, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: {
        legend: {
          position: "bottom",
          labels: { boxWidth: 12, font: { size: 11 } },
        },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}`,
          },
        },
      },
      scales: {
        y: {
          ticks: {
            callback: (v) => v.toLocaleString("ja-JP"),
          },
        },
      },
    },
  });

  // ---------- 表 ----------
  const table = document.getElementById("main-table");
  let html = "<thead><tr><th>年度</th>";
  for (const m of MONTH_LABELS) html += `<th>${m}</th>`;
  html += `<th class="total">${useMode === "cumulative" ? "累計" : "金額"}</th></tr></thead><tbody>`;
  matrix.forEach((row, i) => {
    const values = dataSeries[i];
    const sum = useMode === "cumulative"
      ? values.filter(v => v !== null).slice(-1)[0] ?? null
      : row.values.reduce((a, b) => a + (b || 0), 0);
    const currentClass = i === 0 ? " class=\"current\"" : "";
    html += `<tr${currentClass}><td>${row.display}</td>`;
    for (const v of values) html += `<td>${fmt(v)}</td>`;
    html += `<td class="total">${fmt(sum)}</td></tr>`;
  });
  html += "</tbody>";
  table.innerHTML = html;
}

// ---------- ログイン ----------
async function handleLogin(e) {
  e.preventDefault();
  const id = document.getElementById("login-id").value;
  const pw = document.getElementById("login-pw").value;
  const err = document.getElementById("login-error");
  err.textContent = "";
  const ok = await tryLogin(id, pw);
  if (!ok) {
    err.textContent = "IDまたはパスワードが違います";
    return;
  }
  setAuthed();
  document.getElementById("login-pw").value = "";
  enterAuthed();
}

function enterAuthed() {
  showPage("page-houjin");
}

function logout() {
  clearAuth();
  if (state.chart) { state.chart.destroy(); state.chart = null; }
  showPage("page-login");
}

// ---------- 起動 ----------
async function init() {
  // データ読み込み
  try {
    const res = await fetch(DATA_URL);
    state.data = await res.json();
  } catch (e) {
    alert("データ読み込みに失敗しました: " + e.message);
    return;
  }

  // モード切替
  document.querySelectorAll('input[name="mode"]').forEach(r => {
    r.addEventListener("change", (e) => {
      state.mode = e.target.value;
      if (state.shisetsu && state.koumoku) renderChartAndTable();
    });
  });

  // ログインフォーム
  document.getElementById("login-form").addEventListener("submit", handleLogin);

  // 法人ボタン
  bindHoujinButtons();

  // 認証済みなら法人選択へ
  if (isAuthed()) enterAuthed();
}

init();
