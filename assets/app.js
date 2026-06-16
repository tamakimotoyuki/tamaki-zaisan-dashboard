// 月次経営ダッシュボード メインアプリ
// データ: data/大項目.json
// 構造: sheets[シート名].blocks[項目名][年度ラベル] = {"0": 4月値, ..., "11": 3月値}

const DATA_URL = "data/大項目.json";
const MONTH_LABELS_APR = ["4月","5月","6月","7月","8月","9月","10月","11月","12月","1月","2月","3月"];
const MONTH_LABELS_MAY = ["5月","6月","7月","8月","9月","10月","11月","12月","1月","2月","3月","4月"];

// 年度ラベルから月開始を判定 (4=4月始まり医療法人系, 5=5月始まりMS)
function detectMonthStart(yearLabel) {
  if (/^\d+\.5-\d+\.4$/.test(yearLabel)) return 5;
  return 4;
}
function getMonthLabels(monthStart) {
  return monthStart === 5 ? MONTH_LABELS_MAY : MONTH_LABELS_APR;
}
// シート全体の月開始を判定（最初の年度ラベルから）
function detectSheetMonthStart(sheet) {
  if (!sheet || !sheet.blocks) return 4;
  for (const item of Object.values(sheet.blocks)) {
    const yrs = Object.keys(item);
    if (yrs.length) return detectMonthStart(yrs[0]);
  }
  return 4;
}

// ダッシュボードでのみ使うシート（通常の施設タブからは非表示）
const DASHBOARD_ONLY_SHEETS = new Set([
  "【自動】 統合)明和会+MS",
]);

// 法人 → 所属シート（xlsx 並び順）
const HOUJIN_SHEETS = {
  meiwakai: [
    "【自動】 統合)明和会+MS",
    "医）明和会　全体(PL)",
    "医）明和会　全体(BS)",
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
    "メディエンス　全体(PL)",
    "メディエンス　全体(BS)",
    "メディエンス　ハート徳島クリニック",
  ],
  ms: [
    "MS　全体(PL)",
    "MS　全体(BS)",
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
    "社福　全体(PL)",
    "社福　全体(BS)",
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
  "(水道光熱費)",
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

// ブロックの種別判定（BSタブ整理用）
//  empty   : 全年度・全月が0/空 → 非表示にする
//  annual  : 各年度で「年度末月(=最大月index)」だけに値 → 決算期のみの年間値(CF/純増減系)
//  monthly : それ以外（月次残高あり）
function classifyBlockKind(block) {
  let anyNonZero = false, anyNonYearEnd = false;
  for (const y of Object.keys(block || {})) {
    const months = block[y] || {};
    const idxs = Object.keys(months).filter(k => /^\d+$/.test(k)).map(Number);
    if (!idxs.length) continue;
    const maxIdx = Math.max(...idxs);
    for (const i of idxs) {
      const v = months[String(i)];
      if (typeof v === "number" && v !== 0) {
        anyNonZero = true;
        if (i !== maxIdx) anyNonYearEnd = true;
      }
    }
  }
  if (!anyNonZero) return "empty";
  return anyNonYearEnd ? "monthly" : "annual";
}

// BS月次残高項目の表示順（資産→負債→純資産。リスト外は五十音）
const BS_STOCK_ORDER = [
  "現預金", "医業未収金", "棚卸資産", "有形固定資産（簿価）",
  "未払金・未払費用", "短期借入金", "長期借入金", "純資産",
];
const BS_STOCK_RANK = Object.fromEntries(BS_STOCK_ORDER.map((k, i) => [k, i]));
function bsStockSort(a, b) {
  const ra = BS_STOCK_RANK[a] ?? 999, rb = BS_STOCK_RANK[b] ?? 999;
  if (ra !== rb) return ra - rb;
  return a.localeCompare(b, "ja");
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

// 年度ラベル → 表示用
// 4-3パターン: 医療法人(明和会・メディエンス・社福) - 4月始まり3月終わり → "R7" 等
// 5-4パターン: MS(株式会社) - 5月始まり4月終わり → "FY7" 等
function yearLabelDisplay(label) {
  let m = label.match(/^(\d+)\.4-(\d+)\.3$/);
  if (m) {
    const start = parseInt(m[1]);
    if (start >= 25 && start <= 31) return `H${start}`;
    return `R${start}`;
  }
  m = label.match(/^(\d+)\.5-(\d+)\.4$/);
  if (m) {
    const start = parseInt(m[1]);
    if (start >= 25 && start <= 31) return `H${start}`;
    return `R${start}`;  // MSもR表記で統一（FYと違うので注意）
  }
  return label;
}

// 年度数値（古い→新しいで増える）。並び替え用
function yearOrderKey(label) {
  // 4-3 (医療法人): start=25-31 平成, それ以外は令和
  let m = label.match(/^(\d+)\.4-(\d+)\.3$/);
  if (m) {
    const start = parseInt(m[1]);
    if (start >= 25 && start <= 31) return 1988 + start;
    return 2018 + start;
  }
  // 5-4 (MS) も同じく令和ベース。5月始まりなので少しずらす
  m = label.match(/^(\d+)\.5-(\d+)\.4$/);
  if (m) {
    const start = parseInt(m[1]);
    if (start >= 25 && start <= 31) return 1988 + start + 0.1;
    return 2018 + start + 0.1;  // 同年4-3の直後に並ぶよう微オフセット
  }
  return 0;
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

// ---------- View Transitions ヘルパー ----------
// 画面・タブ・単月/累積の切替を上品なクロスフェードにする。
// rootのみ（view-transition-name を付けない）＝重複リスクなし。グラフは Chart.js animation:false の静止canvas＝フェードで滲まない。
// 非対応ブラウザ／OSの「動きを減らす」設定では updateFn をそのまま実行＝アニメ無しで普通に動く（検出ガード必須）。
// 規約・落とし穴の正本: 【デザイン】/view-transitions.md ／ ui-patterns.md §6-B
// 連続で startViewTransition を呼ぶと前の遷移が中断され InvalidStateError が出る。
// 画面表示前の下準備（オフスクリーン描画）では遷移を抑止し、見える切替だけを1回アニメさせる。
let suppressVT = false;
function withViewTransition(updateFn) {
  const reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (document.startViewTransition && !reduce && !suppressVT) {
    const t = document.startViewTransition(updateFn);
    // 遷移は“装飾”として扱う：中断（連続呼び出しでの abort）・name重複などで ready/finished が
    // reject してもアプリに影響させない（DOM更新=updateFn は既に完了している）。
    // これを握りつぶさないと InvalidStateError 等が未処理 rejection としてコンソールに出る。
    if (t) {
      if (t.ready) t.ready.catch(() => {});
      if (t.finished) t.finished.catch(() => {});
    }
  } else {
    updateFn();
  }
}

// ---------- 画面切替 ----------
function showPage(id) {
  withViewTransition(() => {
    document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
    document.getElementById(id).classList.add("active");
  });
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
  // ダッシュボード起動
  document.getElementById("open-dashboard").addEventListener("click", () => {
    renderGlobalDashboard();
    showPage("page-global-dashboard");
  });
  // ダッシュボード単月/累積トグル
  document.querySelectorAll('input[name="dashboard-mode"]').forEach(r => {
    r.addEventListener("change", (e) => {
      dashboardMode = e.target.value;
      withViewTransition(renderGlobalDashboard);
    });
  });
  document.getElementById("back-to-houjin-from-dashboard").addEventListener("click", () => {
    destroyDashboardCharts();
    showPage("page-houjin");
  });
  document.getElementById("logout-dashboard").addEventListener("click", () => {
    destroyDashboardCharts();
    logout();
  });
}

function enterHoujin(houjin) {
  state.houjin = houjin;
  state.shisetsu = null;
  state.koumoku = null;
  document.getElementById("houjin-name").textContent = HOUJIN_LABELS[houjin];
  // 表示前の下準備は遷移させない（showPage の1回だけアニメ）＝二重遷移による中断を防ぐ
  suppressVT = true;
  renderShisetsuTabs();
  // 最初の施設を選択
  const sheets = HOUJIN_SHEETS[houjin].filter(s => state.data.sheets[s] && !DASHBOARD_ONLY_SHEETS.has(s));
  if (sheets.length) selectShisetsu(sheets[0]);
  suppressVT = false;
  showPage("page-dashboard");
}

function renderShisetsuTabs() {
  const wrap = document.getElementById("shisetsu-tabs");
  wrap.innerHTML = "";
  const sheets = HOUJIN_SHEETS[state.houjin].filter(s => state.data.sheets[s] && !DASHBOARD_ONLY_SHEETS.has(s));
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

// BSシート判定 (全体(BS) 等)
function isBalanceSheet(sheetName) {
  return /\(BS\)\s*$/.test(sheetName || "");
}

function selectShisetsu(sheetName) {
  state.shisetsu = sheetName;
  document.querySelectorAll("#shisetsu-tabs .tab").forEach(t => {
    t.classList.toggle("active", t.dataset.sheet === sheetName);
  });
  // BSタブでは「単月/累積」トグル非表示・モード固定(monthly=月末残高そのまま)
  // (BSデータは既に月末累計残高=ストック値・累積化は無意味)
  const toggleWrap = document.querySelector("#page-dashboard .chart-toggle-wrap");
  if (toggleWrap) {
    if (isBalanceSheet(sheetName)) {
      toggleWrap.style.display = "none";
      state.mode = "monthly";  // 強制リセット
    } else {
      toggleWrap.style.display = "";
    }
  }
  // ★トグルradioの見た目を state.mode に同期（施設切替=特にBS→PLでのズレ防止・2026-05-28）
  const modeRadio = document.querySelector(`#page-dashboard input[name="mode"][value="${state.mode}"]`);
  if (modeRadio) modeRadio.checked = true;
  renderKoumokuTabs();
  // 並べ替え後の最初のタブを選択（BSは月次残高が先頭に来る）
  const firstTab = document.querySelector("#koumoku-tabs .tab");
  if (firstTab) selectKoumoku(firstTab.dataset.item);
}

function renderKoumokuTabs() {
  const wrap = document.getElementById("koumoku-tabs");
  wrap.innerHTML = "";
  const sheet = state.data.sheets[state.shisetsu];
  const blockKeys = Object.keys(sheet.blocks);

  const mkTab = (k, isAnnual) => {
    const btn = document.createElement("button");
    btn.className = "tab" + (isAnnual ? " tab-annual" : "");
    btn.textContent = isAnnual ? displayItemName(k) + "（年間）" : displayItemName(k);
    btn.title = isAnnual ? k + "（決算期=年度末のみの年間値・CF計算用）" : k;
    btn.dataset.item = k;
    btn.dataset.kind = isAnnual ? "annual" : "monthly";
    btn.addEventListener("click", () => selectKoumoku(k));
    wrap.appendChild(btn);
  };

  if (isBalanceSheet(state.shisetsu)) {
    // BS: 空項目は非表示。月次残高(左)→年度末のみの年間/CF項目(右)に分ける
    const monthly = [], annual = [];
    for (const k of blockKeys) {
      const kind = classifyBlockKind(sheet.blocks[k]);
      if (kind === "empty") continue;
      (kind === "annual" ? annual : monthly).push(k);
    }
    monthly.sort(bsStockSort).forEach(k => mkTab(k, false));
    if (annual.length) {
      const sep = document.createElement("span");
      sep.className = "tab-group-sep";
      sep.textContent = "｜決算期のみ▶";
      wrap.appendChild(sep);
      annual.sort((a, b) => a.localeCompare(b, "ja")).forEach(k => mkTab(k, true));
    }
    return;
  }
  // 非BS（PL等）は従来どおり
  sortItems(blockKeys).forEach(k => mkTab(k, false));
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
  withViewTransition(renderChartAndTableBody);
}
function renderChartAndTableBody() {
  // BSの年度末のみ項目（CF/純増減系）は月次でなく年度末値の年次推移で表示
  if (isBalanceSheet(state.shisetsu)) {
    const blk = state.data.sheets[state.shisetsu].blocks[state.koumoku];
    if (classifyBlockKind(blk) === "annual") { renderAnnualItem(blk); return; }
  }
  const matrix = getMatrix();
  const useMode = state.mode;
  const dataSeries = matrix.map(row => useMode === "cumulative" ? cumulative(row.values) : row.values);

  // 月ラベル（シートの年度ラベルから5月始まり/4月始まりを判定）
  const monthStart = detectSheetMonthStart(state.data.sheets[state.shisetsu]);
  const monthLabels = getMonthLabels(monthStart);

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
    data: { labels: monthLabels, datasets },
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
  for (const m of monthLabels) html += `<th>${m}</th>`;
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

// ---------- BS年度末のみ項目（CF/純増減系）の年次推移表示 ----------
// 月次を持たない項目は、各年度の年度末値だけを年次バーで見せる（「4月しかない」誤読を防ぐ）
function renderAnnualItem(block) {
  // 古い→新しい順に年度末値を並べる
  const rows = Object.keys(block)
    .sort((a, b) => yearOrderKey(a) - yearOrderKey(b))
    .map(y => ({ display: yearLabelDisplay(y), value: yearEndFromBlock(block, y) }))
    .filter(r => r.value !== null);

  const ctx = document.getElementById("main-chart").getContext("2d");
  if (state.chart) state.chart.destroy();
  const total = rows.length;
  state.chart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: rows.map(r => r.display),
      datasets: [{
        label: "決算期（年度末）",
        data: rows.map(r => r.value),
        backgroundColor: rows.map((_, i) => yearColor(total - 1 - i, total)),
        borderColor: rows.map((_, i) => yearColor(total - 1 - i, total)),
        borderWidth: 1,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => `決算期: ${fmt(ctx.parsed.y)}` } },
      },
      scales: { y: { ticks: { callback: (v) => v.toLocaleString("ja-JP") } } },
    },
  });

  // 表（新→古）
  const table = document.getElementById("main-table");
  let html = "<thead><tr><th>年度</th><th class=\"total\">決算期（年度末）金額</th></tr></thead><tbody>";
  rows.slice().reverse().forEach((r, i) => {
    const cls = i === 0 ? " class=\"current\"" : "";
    html += `<tr${cls}><td>${r.display}</td><td class="total">${fmt(r.value)}</td></tr>`;
  });
  html += "</tbody>";
  table.innerHTML = html;
}

// ---------- 横断ダッシュボード ----------
// 人件費率(給与費÷収益)の分母候補（法人で収益科目名が違うので候補から自動選択）
const RATIO_DEN = ["売上高", "純売上高", "事業収益合計", "サービス活動収益", "医業収益（全体）", "医業収益", "医療収益"];
const DASHBOARD_SECTIONS = [
  {
    title: "①【統合】明和会＋MS",
    sheet: "【自動】 統合)明和会+MS",
    items: {
      "経常利益": ["経常利益（明和会+MS）", "経常利益", "経常損益"],
      "医業収益": ["明和会医業収入＋MS 売上高", "医業収益（全体）", "医業収益", "事業収益合計", "売上高"],
    },
  },
  {
    title: "②（医）明和会",
    sheet: "医）明和会　全体(PL)",
    items: {
      "経常利益": ["経常利益", "経常損益"],
      "医療収益": ["医業収益（全体）", "医業収益", "医療収益", "事業収益合計"],
      "人件費率(%)": { ratio: true, num: ["給与費合計"], den: RATIO_DEN },
    },
  },
  {
    title: "③（有）たまきメディカルサポート",
    sheet: "MS　全体(PL)",
    items: {
      "経常利益": ["経常利益", "経常損益"],
      "売上高":   ["売上高", "純売上高"],
      "人件費率(%)": { ratio: true, num: ["給与費合計"], den: RATIO_DEN },
    },
  },
  {
    title: "④（医）メディエンス",
    sheet: "メディエンス　全体(PL)",
    items: {
      "経常利益": ["経常損益", "経常利益"],
      "事業収益": ["事業収益合計", "医業収益", "医療収益"],
      "人件費率(%)": { ratio: true, num: ["給与費合計"], den: RATIO_DEN },
    },
  },
  {
    title: "⑤（社福）明和福祉会",
    sheet: "社福　全体(PL)",
    items: {
      "経常利益": ["経常利益", "経常損益"],
      "サービス活動収益": ["サービス活動収益", "事業収益合計"],
      "人件費率(%)": { ratio: true, num: ["給与費合計"], den: RATIO_DEN },
    },
  },
  // 純資産推移 - 法人別4チャート分割（最重要KPI）
  {
    title: "⑥純資産推移",
    sub: "4法人別BS・年度末残高",
    sheet: null,
    isBS: true,
    bsSheets: {
      "（医）明和会": "医）明和会　全体(BS)",
      "（医）メディエンス": "メディエンス　全体(BS)",
      "MS": "MS　全体(BS)",
      "（社福）明和福祉会": "社福　全体(BS)",
    },
    items: {
      "純資産": "純資産",
    },
    splitByHoujin: ["純資産"],
    minYear: "R2",  // 全法人R2スタートに統一
  },
  // 現預金推移 - 法人別4チャート分割（流動性KPI）
  {
    title: "⑦現預金推移",
    sub: "4法人別BS・年度末残高",
    sheet: null,
    isBS: true,
    bsSheets: {
      "（医）明和会": "医）明和会　全体(BS)",
      "（医）メディエンス": "メディエンス　全体(BS)",
      "MS": "MS　全体(BS)",
      "（社福）明和福祉会": "社福　全体(BS)",
    },
    items: {
      "現預金": "現預金",
    },
    splitByHoujin: ["現預金"],
    minYear: "R2",  // 全法人R2スタートに統一
  },
  // 長期借入金推移 - 法人別4チャート分割（データ欠損は0扱い）
  {
    title: "⑧長期借入金推移",
    sub: "4法人別BS・年度末残高",
    sheet: null,
    isBS: true,
    bsSheets: {
      "（医）明和会": "医）明和会　全体(BS)",
      "（医）メディエンス": "メディエンス　全体(BS)",
      "MS": "MS　全体(BS)",
      "（社福）明和福祉会": "社福　全体(BS)",
    },
    items: {
      "長期借入金": "長期借入金",
    },
    splitByHoujin: ["長期借入金"],
    fillMissingAsZero: ["長期借入金"],
    barChart: true,   // 棒グラフ
    yMin0: true,      // Y軸0底辺(借入金は非負)
    minYear: "R2",    // R2以降に統一
  },
  // 有形固定資産推移 - 法人別4チャート分割（設備投資の累計）
  {
    title: "⑨有形固定資産推移",
    sub: "4法人別BS・年度末残高",
    sheet: null,
    isBS: true,
    bsSheets: {
      "（医）明和会": "医）明和会　全体(BS)",
      "（医）メディエンス": "メディエンス　全体(BS)",
      "MS": "MS　全体(BS)",
      "（社福）明和福祉会": "社福　全体(BS)",
    },
    items: {
      "有形固定資産（簿価）": "有形固定資産（簿価）",
    },
    splitByHoujin: ["有形固定資産（簿価）"],
    barChart: true,   // 棒グラフ
    yMin0: true,      // Y軸0底辺(簿価は非負)
    minYear: "R2",    // R2以降に統一
  },
  // 減価償却費推移 - 法人別4チャート分割（PL年度合計）
  {
    title: "⑩減価償却費推移",
    sub: "4法人別PL・月次（年合計÷12）",
    isPL: true,
    houjinSheets: {
      "（医）明和会": "医）明和会　全体(PL)",
      "（医）メディエンス": "メディエンス　全体(PL)",
      "MS": "MS　全体(PL)",
      "（社福）明和福祉会": "社福　全体(PL)",
    },
    items: {
      "減価償却費": ["減価償却費"],
    },
    minYear: "R2",  // R2以降に統一
  },
  // 簡易営業CF (経常利益 + 減価償却費) - 法人別4チャート・スタック
  {
    title: "⑪簡易営業CF推移",
    sub: "経常利益＋減価償却費・月次・年度比較",
    isStackedPL: true,
    houjinSheets: {
      "（医）明和会": "医）明和会　全体(PL)",
      "（医）メディエンス": "メディエンス　全体(PL)",
      "MS": "MS　全体(PL)",
      "（社福）明和福祉会": "社福　全体(PL)",
    },
    stackItems: [
      { label: "経常利益", candidates: ["経常利益", "経常損益"] },
      { label: "減価償却費", candidates: ["減価償却費"], evenDistribute: true },
    ],
    minYear: "R2",  // R2以降に統一
  },
  // 簡易キャッシュフロー (営業/投資/財務) - 法人別4チャート
  {
    title: "⑫簡易キャッシュフロー推移",
    sub: "営業/投資/財務・年度概算",
    isCF: true,
    minYear: "R2",  // R2以降に統一
    houjinPlSheets: {
      "（医）明和会": "医）明和会　全体(PL)",
      "（医）メディエンス": "メディエンス　全体(PL)",
      "MS": "MS　全体(PL)",
      "（社福）明和福祉会": "社福　全体(PL)",
    },
    houjinBsSheets: {
      "（医）明和会": "医）明和会　全体(BS)",
      "（医）メディエンス": "メディエンス　全体(BS)",
      "MS": "MS　全体(BS)",
      "（社福）明和福祉会": "社福　全体(BS)",
    },
  },
];

let dashboardMode = "monthly"; // monthly | cumulative

// セクション見出し：短い<h2>＋詳細は小字サブ（projector/印刷でも読めるよう常時表示・控えめ）
function appendSectionHeader(sec, section) {
  const h = document.createElement("h2");
  h.textContent = section.title;
  sec.appendChild(h);
  if (section.sub) {
    const s = document.createElement("div");
    s.className = "dash-sub";
    s.textContent = section.sub;
    sec.appendChild(s);
  }
}

const dashboardCharts = [];  // Chart instance管理

function findFirstItem(blocks, candidates) {
  for (const k of candidates) if (blocks[k]) return k;
  return null;
}

function destroyDashboardCharts() {
  dashboardCharts.forEach(c => c && c.destroy());
  dashboardCharts.length = 0;
}

function renderGlobalDashboard() {
  const grid = document.getElementById("dashboard-grid");
  grid.innerHTML = "";
  destroyDashboardCharts();

  for (const section of DASHBOARD_SECTIONS) {
    // BSセクション専用処理 (4法人横断・純資産/現預金推移)
    if (section.isBS) {
      renderBSDashboardSection(grid, section);
      continue;
    }
    if (section.isPL) {
      renderPLAnnualSection(grid, section);
      continue;
    }
    if (section.isStackedPL) {
      renderStackedPLSection(grid, section);
      continue;
    }
    if (section.isCF) {
      renderCFSection(grid, section);
      continue;
    }
    if (section.isBorrowFlow) {
      renderBorrowFlowSection(grid, section);
      continue;
    }
    const sheet = state.data.sheets[section.sheet];
    const sec = document.createElement("section");
    sec.className = "dashboard-section";
    appendSectionHeader(sec, section);
    const chartsDiv = document.createElement("div");
    chartsDiv.className = "dashboard-charts";
    sec.appendChild(chartsDiv);

    for (const [label, candidates] of Object.entries(section.items)) {
      const cdiv = document.createElement("div");
      cdiv.className = "dashboard-chart";
      const t = document.createElement("div");
      t.className = "dashboard-chart-title";
      t.textContent = label;
      cdiv.appendChild(t);

      if (!sheet) {
        const e = document.createElement("div");
        e.className = "dashboard-chart-empty";
        e.textContent = "シートなし";
        cdiv.appendChild(e);
        chartsDiv.appendChild(cdiv);
        continue;
      }
      // ratio型(人件費率など%指標) か 配列(金額) かで分岐
      const isRatio = candidates && !Array.isArray(candidates) && candidates.ratio;
      let years, datasets;
      let isPercent = false;

      if (isRatio) {
        const numKey = findFirstItem(sheet.blocks, candidates.num);
        const denKey = findFirstItem(sheet.blocks, candidates.den);
        if (!numKey || !denKey) {
          const e = document.createElement("div");
          e.className = "dashboard-chart-empty";
          e.textContent = "データなし";
          cdiv.appendChild(e);
          chartsDiv.appendChild(cdiv);
          continue;
        }
        isPercent = true;
        const nb = sheet.blocks[numKey], db = sheet.blocks[denKey];
        years = Object.keys(nb).filter(y => db[y]).slice().sort((a, b) => yearOrderKey(b) - yearOrderKey(a));
        const total = years.length;
        // 人件費率: 単月=月給与費/月収益, 累計=累計給与費/累計収益 (ユーザー方針2026-05-27)
        datasets = years.map((y, idx) => {
          const na = monthArray(nb[y]), da = monthArray(db[y]);
          let data;
          if (dashboardMode === "cumulative") {
            const cn = cumulative(na), cd = cumulative(da);
            data = cn.map((v, i) => (cd[i] ? +(v / cd[i] * 100).toFixed(1) : null));
          } else {
            data = na.map((v, i) => (da[i] ? +(v / da[i] * 100).toFixed(1) : null));
          }
          return {
            label: yearLabelDisplay(y), data,
            backgroundColor: yearColor(idx, total), borderColor: yearColor(idx, total), borderWidth: 1,
          };
        });
      } else {
        const itemKey = findFirstItem(sheet.blocks, candidates);
        if (!itemKey) {
          const e = document.createElement("div");
          e.className = "dashboard-chart-empty";
          e.textContent = "データなし";
          cdiv.appendChild(e);
          chartsDiv.appendChild(cdiv);
          continue;
        }
        const block = sheet.blocks[itemKey];
        years = Object.keys(block).slice().sort((a, b) => yearOrderKey(b) - yearOrderKey(a));
        const total = years.length;
        datasets = years.map((y, idx) => ({
          label: yearLabelDisplay(y),
          data: dashboardMode === "cumulative" ? cumulative(monthArray(block[y])) : monthArray(block[y]),
          backgroundColor: yearColor(idx, total),
          borderColor: yearColor(idx, total),
          borderWidth: 1,
        }));
      }

      const wrap = document.createElement("div");
      wrap.className = "dashboard-chart-canvas";
      const canvas = document.createElement("canvas");
      wrap.appendChild(canvas);
      cdiv.appendChild(wrap);
      chartsDiv.appendChild(cdiv);

      // 月ラベル: このシート/グラフ用に判定
      const sheetMonthStart = detectSheetMonthStart(sheet);
      const sheetMonthLabels = getMonthLabels(sheetMonthStart);
      const c = new Chart(canvas.getContext("2d"), {
        type: "bar",
        data: { labels: sheetMonthLabels, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: { label: (ctx) => isPercent ? `${ctx.dataset.label}: ${ctx.parsed.y}%` : `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}` },
            },
          },
          scales: {
            x: { ticks: { font: { size: 9 } } },
            // 人件費率(%)はY軸下限も自動（データmin付近）＝変動を拡大表示。金額は0始まり
            y: { beginAtZero: !isPercent, ticks: { font: { size: 9 }, callback: v => isPercent ? v + "%" : v.toLocaleString("ja-JP") } },
          },
        },
      });
      dashboardCharts.push(c);
    }
    grid.appendChild(sec);
  }
}

// ---------- BSダッシュボードセクション (4法人横断) ----------
// 各項目 (純資産/現預金) について、4法人の年度末残高を1グラフに重ね描き。
// BSは月末残高=ストック値なので、各年度の3月末 (=最終月) の値を取って年度推移を可視化。
function renderBSDashboardSection(grid, section) {
  const sec = document.createElement("section");
  sec.className = "dashboard-section";
  appendSectionHeader(sec, section);
  const chartsDiv = document.createElement("div");
  chartsDiv.className = "dashboard-charts";
  sec.appendChild(chartsDiv);

  const splitItems = new Set(section.splitByHoujin || []);
  const fillZeroItems = new Set(section.fillMissingAsZero || []);

  // 法人ごとの「他BS項目で利用可能な年度」をプリ計算（欠損→0補完用）
  const houjinAvailableYears = {}; // houjinLabel -> Set<yearLabel>
  for (const [houjinLabel, sheetName] of Object.entries(section.bsSheets)) {
    const sheet = state.data.sheets[sheetName];
    const ys = new Set();
    if (sheet && sheet.blocks) {
      for (const blockKey of Object.keys(sheet.blocks)) {
        const block = sheet.blocks[blockKey] || {};
        for (const yearLabel of Object.keys(block)) {
          const months = block[yearLabel] || {};
          const monthKeys = Object.keys(months).filter(k => /^\d+$/.test(k));
          if (monthKeys.length > 0) ys.add(yearLabel);
        }
      }
    }
    houjinAvailableYears[houjinLabel] = ys;
  }

  // 各項目について4法人重ね折れ線グラフ
  for (const [label, itemKey] of Object.entries(section.items)) {
    const fillZero = fillZeroItems.has(label);
    // 全法人で利用可能な年度ラベルを収集
    const allYears = new Set();
    const houjinData = {}; // houjinLabel -> {year -> value (年度末月)}
    for (const [houjinLabel, sheetName] of Object.entries(section.bsSheets)) {
      const sheet = state.data.sheets[sheetName];
      houjinData[houjinLabel] = {};
      const block = sheet?.blocks?.[itemKey] || {};
      for (const yearLabel of Object.keys(block)) {
        const months = block[yearLabel] || {};
        const monthKeys = Object.keys(months).filter(k => /^\d+$/.test(k));
        if (monthKeys.length === 0) continue;
        const maxKey = Math.max(...monthKeys.map(Number));
        const yearEnd = months[String(maxKey)];
        if (typeof yearEnd === "number") {
          houjinData[houjinLabel][yearLabel] = yearEnd;
          allYears.add(yearLabel);
        }
      }
      // 欠損→0補完: その法人が他BS項目で持っている年度に値がなければ0扱い
      if (fillZero) {
        for (const y of houjinAvailableYears[houjinLabel]) {
          if (!(y in houjinData[houjinLabel])) {
            houjinData[houjinLabel][y] = 0;
            allYears.add(y);
          }
        }
      }
    }

    const sortedYears = dedupeYearsByDisplay(Array.from(allYears)).sort((a, b) => yearOrderKey(a) - yearOrderKey(b));
    const yearDisplays = sortedYears.map(y => yearLabelDisplay(y));
    const houjinLabels = Object.keys(section.bsSheets);

    if (splitItems.has(label)) {
      // 法人別4チャート分割 — 全法人で共通の表示軸(R2..最新)を使い、横軸を揃える
      // 各法人は「displayラベル→値」で引く (法人間のキー形式 4月/5月始まり 混在を吸収)
      const minKey = section.minYear
        ? yearOrderKey(section.minYear === "R2" ? "2.4-3.3" : section.minYear)
        : -Infinity;
      const dispOrder = new Map();  // display -> orderKey
      houjinLabels.forEach(hLabel => {
        Object.keys(houjinData[hLabel] || {}).forEach(y => {
          if (yearOrderKey(y) >= minKey) {
            const disp = yearLabelDisplay(y);
            if (!dispOrder.has(disp)) dispOrder.set(disp, yearOrderKey(y));
          }
        });
      });
      const commonDisplays = [...dispOrder.keys()].sort((a, b) => dispOrder.get(a) - dispOrder.get(b));

      houjinLabels.forEach((hLabel, idx) => {
        const cdiv = document.createElement("div");
        cdiv.className = "dashboard-chart";
        const t = document.createElement("div");
        t.className = "dashboard-chart-title";
        t.textContent = `${label} - ${hLabel}（年度末月末残高）`;
        cdiv.appendChild(t);
        const wrap = document.createElement("div");
        wrap.className = "dashboard-chart-canvas";
        const canvas = document.createElement("canvas");
        wrap.appendChild(canvas);
        cdiv.appendChild(wrap);
        chartsDiv.appendChild(cdiv);

        // この法人の display→値 マップを作り、共通軸に沿って値を引く
        const dispToVal = {};
        Object.keys(houjinData[hLabel] || {}).forEach(y => {
          dispToVal[yearLabelDisplay(y)] = houjinData[hLabel][y];
        });
        const myDisplays = commonDisplays;
        const data = commonDisplays.map(disp => (disp in dispToVal ? dispToVal[disp] : null));

        if (data.every(v => v === null)) {
          const e = document.createElement("div");
          e.className = "dashboard-chart-empty";
          e.textContent = "BSデータなし";
          cdiv.appendChild(e);
          return;
        }
        const color = yearColor(idx, houjinLabels.length);
        const isBar = !!section.barChart;
        const c = new Chart(canvas.getContext("2d"), {
          type: isBar ? "bar" : "line",
          data: {
            labels: myDisplays,
            datasets: [{
              label: hLabel,
              data,
              backgroundColor: color,
              borderColor: color,
              borderWidth: isBar ? 1 : 2,
              tension: 0.1,
              spanGaps: true,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            plugins: {
              legend: { display: false },
              tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}` } },
            },
            scales: {
              x: { ticks: { font: { size: 10 } } },
              y: { beginAtZero: !!section.yMin0, min: section.yMin0 ? 0 : undefined,
                   ticks: { font: { size: 9 }, callback: v => v.toLocaleString("ja-JP") } },
            },
          },
        });
        dashboardCharts.push(c);
      });
      continue;
    }

    // 4法人重ね（既定）
    const cdiv = document.createElement("div");
    cdiv.className = "dashboard-chart";
    const t = document.createElement("div");
    t.className = "dashboard-chart-title";
    t.textContent = label + "(年度末月末残高)";
    cdiv.appendChild(t);
    const wrap = document.createElement("div");
    wrap.className = "dashboard-chart-canvas";
    const canvas = document.createElement("canvas");
    wrap.appendChild(canvas);
    cdiv.appendChild(wrap);
    chartsDiv.appendChild(cdiv);

    const datasets = houjinLabels.map((hLabel, idx) => ({
      label: hLabel,
      data: sortedYears.map(y => houjinData[hLabel]?.[y] ?? null),
      backgroundColor: yearColor(idx, houjinLabels.length),
      borderColor: yearColor(idx, houjinLabels.length),
      borderWidth: 2,
      tension: 0.1,
      spanGaps: true,
    }));

    if (datasets.every(d => d.data.every(v => v === null))) {
      const e = document.createElement("div");
      e.className = "dashboard-chart-empty";
      e.textContent = "BSデータなし";
      cdiv.appendChild(e);
      continue;
    }

    const c = new Chart(canvas.getContext("2d"), {
      type: "line",
      data: { labels: yearDisplays, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        plugins: {
          legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}` } },
        },
        scales: {
          x: { ticks: { font: { size: 10 } } },
          y: { ticks: { font: { size: 9 }, callback: v => v.toLocaleString("ja-JP") } },
        },
      },
    });
    dashboardCharts.push(c);
  }
  grid.appendChild(sec);
}

// ---------- 同一表示ラベル年度の重複排除 (MS PL の 7.4-8.3 + 7.5-8.4 等の保険) ----------
function dedupeYearsByDisplay(years) {
  const seen = new Map();
  for (const y of years) {
    const disp = yearLabelDisplay(y);
    // 同じ表示なら文字列比較で大きい方 (5月始まり .5- > 4月始まり .4-) を採用
    if (!seen.has(disp) || y > seen.get(disp)) seen.set(disp, y);
  }
  return years.filter(y => seen.get(yearLabelDisplay(y)) === y);
}

// ---------- PL年度合計ヘルパー ----------
function annualSumFromBlock(block, year) {
  const months = block?.[year] || {};
  const keys = Object.keys(months).filter(k => /^\d+$/.test(k));
  if (keys.length === 0) return null;
  let sum = 0;
  let any = false;
  for (const k of keys) {
    if (typeof months[k] === "number") { sum += months[k]; any = true; }
  }
  return any ? sum : null;
}
// 減価償却の月次配列を返す (ハイブリッド)
// ・実際に月次入力されている年度 → 生の月次値をそのまま
// ・年度末に一括計上の年度 (旧年度) → 年合計÷12 で均す
// 判定: 最終月の絶対値が「全月絶対値合計」の50%以上を占める or 月数1 → 一括 → ÷12
function depreciationMonthlyArray(block, year) {
  const months = block?.[year] || {};
  const arr = new Array(12).fill(null);
  let any = false, sumAbs = 0, lastAbs = 0, annual = 0, nonZero = 0;
  for (let m = 0; m < 12; m++) {
    const v = months[String(m)];
    if (typeof v === "number") {
      arr[m] = v; any = true; annual += v; sumAbs += Math.abs(v);
      if (Math.abs(v) > 0) nonZero++;
      if (m === 11) lastAbs = Math.abs(v);
    }
  }
  if (!any) return arr;
  const isLump = sumAbs > 0 && (nonZero <= 1 || (lastAbs / sumAbs) >= 0.5);
  if (isLump) return new Array(12).fill(annual / 12);  // 旧年度: ÷12で均す
  return arr;  // 新年度: 実月次
}
function yearEndFromBlock(block, year) {
  const months = block?.[year] || {};
  const keys = Object.keys(months).filter(k => /^\d+$/.test(k));
  if (keys.length === 0) return null;
  const maxKey = Math.max(...keys.map(Number));
  const v = months[String(maxKey)];
  return typeof v === "number" ? v : null;
}
function extractHoujinAnnualSum(houjinSheets, candidates) {
  // returns {houjinLabel -> {year -> annualSum}}, allYears Set
  const result = {};
  const allYears = new Set();
  for (const [hLabel, sheetName] of Object.entries(houjinSheets)) {
    const sheet = state.data.sheets[sheetName];
    result[hLabel] = {};
    if (!sheet?.blocks) continue;
    const key = findFirstItem(sheet.blocks, candidates);
    if (!key) continue;
    const block = sheet.blocks[key];
    for (const y of Object.keys(block)) {
      const v = annualSumFromBlock(block, y);
      if (v !== null) { result[hLabel][y] = v; allYears.add(y); }
    }
  }
  return { result, allYears };
}
function extractHoujinYearEnd(bsSheets, itemKey) {
  const result = {};
  const allYears = new Set();
  for (const [hLabel, sheetName] of Object.entries(bsSheets)) {
    const sheet = state.data.sheets[sheetName];
    result[hLabel] = {};
    if (!sheet?.blocks?.[itemKey]) continue;
    const block = sheet.blocks[itemKey];
    for (const y of Object.keys(block)) {
      const v = yearEndFromBlock(block, y);
      if (v !== null) { result[hLabel][y] = v; allYears.add(y); }
    }
  }
  return { result, allYears };
}

// ---------- ⑩ PL長期月次タイムライン (法人別4チャート・年合計÷12で月次均等配分) ----------
// PCAの記帳パターン差を吸収し、長期推移を月次解像度で見られるように
function renderPLAnnualSection(grid, section) {
  const sec = document.createElement("section");
  sec.className = "dashboard-section";
  appendSectionHeader(sec, section);
  const chartsDiv = document.createElement("div");
  chartsDiv.className = "dashboard-charts";
  sec.appendChild(chartsDiv);

  const houjinLabels = Object.keys(section.houjinSheets);
  for (const [label, candidates] of Object.entries(section.items)) {
    houjinLabels.forEach((hLabel, idx) => {
      const sheetName = section.houjinSheets[hLabel];
      const sheet = state.data.sheets[sheetName];

      const cdiv = document.createElement("div");
      cdiv.className = "dashboard-chart";
      const t = document.createElement("div");
      t.className = "dashboard-chart-title";
      cdiv.appendChild(t);
      const wrap = document.createElement("div");
      wrap.className = "dashboard-chart-canvas";
      const canvas = document.createElement("canvas");
      wrap.appendChild(canvas);
      cdiv.appendChild(wrap);
      chartsDiv.appendChild(cdiv);

      if (!sheet?.blocks) {
        t.textContent = `${label} - ${hLabel}`;
        const e = document.createElement("div");
        e.className = "dashboard-chart-empty";
        e.textContent = "シートなし";
        cdiv.appendChild(e);
        return;
      }
      const itemKey = findFirstItem(sheet.blocks, candidates);
      if (!itemKey) {
        t.textContent = `${label} - ${hLabel}`;
        const e = document.createElement("div");
        e.className = "dashboard-chart-empty";
        e.textContent = "PLデータなし";
        cdiv.appendChild(e);
        return;
      }
      const block = sheet.blocks[itemKey];
      let years = dedupeYearsByDisplay(Object.keys(block))
        .slice().sort((a, b) => yearOrderKey(a) - yearOrderKey(b));
      if (section.minYear) {
        const minKey = yearOrderKey(section.minYear === "R2" ? "2.4-3.3" : section.minYear);
        years = years.filter(y => yearOrderKey(y) >= minKey);
      }

      if (years.length === 0) {
        t.textContent = `${label} - ${hLabel}`;
        const e = document.createElement("div");
        e.className = "dashboard-chart-empty";
        e.textContent = "PLデータなし";
        cdiv.appendChild(e);
        return;
      }

      t.textContent = `${label} - ${hLabel}（月次・${yearLabelDisplay(years[0])}〜${yearLabelDisplay(years[years.length - 1])}・新年度=実月次/旧年度=÷12）`;

      const monthsTpl = getMonthLabels(detectSheetMonthStart(sheet));
      const xLabels = [];
      const data = [];
      years.forEach(y => {
        const yDisp = yearLabelDisplay(y);
        const mvals = depreciationMonthlyArray(block, y);  // 実月次 or ÷12 のハイブリッド
        // 累積モード時は年度内累積 (年度をまたぐとリセット)
        let acc = 0;
        for (let m = 0; m < 12; m++) {
          xLabels.push(`${yDisp} ${monthsTpl[m]}`);
          const v = mvals[m];
          if (dashboardMode === "cumulative") {
            if (v === null) data.push(null);
            else { acc += v; data.push(acc); }
          } else {
            data.push(v);
          }
        }
      });

      const color = yearColor(idx, houjinLabels.length);
      const c = new Chart(canvas.getContext("2d"), {
        type: "bar",
        data: {
          labels: xLabels,
          datasets: [{ label: hLabel, data, backgroundColor: color, borderColor: color, borderWidth: 1 }],
        },
        options: {
          responsive: true, maintainAspectRatio: false, animation: false,
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}` } },
          },
          scales: {
            x: { ticks: { font: { size: 8 }, maxRotation: 90, minRotation: 60, autoSkip: true, maxTicksLimit: 30 } },
            y: { ticks: { font: { size: 9 }, callback: v => v.toLocaleString("ja-JP") } },
          },
        },
      });
      dashboardCharts.push(c);
    });
  }
  grid.appendChild(sec);
}

// ---------- ⑪ 経常利益+減価償却費 スタック (法人別4チャート) ----------
function renderStackedPLSection(grid, section) {
  const sec = document.createElement("section");
  sec.className = "dashboard-section";
  appendSectionHeader(sec, section);
  const chartsDiv = document.createElement("div");
  chartsDiv.className = "dashboard-charts";
  sec.appendChild(chartsDiv);

  const houjinLabels = Object.keys(section.houjinSheets);
  houjinLabels.forEach((hLabel) => {
    const sheetName = section.houjinSheets[hLabel];
    const sheet = state.data.sheets[sheetName];

    const cdiv = document.createElement("div");
    cdiv.className = "dashboard-chart";
    const t = document.createElement("div");
    t.className = "dashboard-chart-title";
    cdiv.appendChild(t);
    const wrap = document.createElement("div");
    wrap.className = "dashboard-chart-canvas";
    const canvas = document.createElement("canvas");
    wrap.appendChild(canvas);
    cdiv.appendChild(wrap);
    chartsDiv.appendChild(cdiv);

    if (!sheet?.blocks) {
      t.textContent = `${hLabel} - 経常利益+減価償却費（月次）`;
      const e = document.createElement("div");
      e.className = "dashboard-chart-empty";
      e.textContent = "シートなし";
      cdiv.appendChild(e);
      return;
    }

    const itemBlocks = section.stackItems.map(item => {
      const key = findFirstItem(sheet.blocks, item.candidates);
      return key ? sheet.blocks[key] : null;
    });
    // 経常利益(必須)のある年度 (新→旧で年度色分け)
    const baseBlock = itemBlocks[0];
    let years = baseBlock
      ? dedupeYearsByDisplay(Object.keys(baseBlock)).slice().sort((a, b) => yearOrderKey(b) - yearOrderKey(a))
      : [];
    if (section.minYear) {
      const minKey = yearOrderKey(section.minYear === "R2" ? "2.4-3.3" : section.minYear);
      years = years.filter(y => yearOrderKey(y) >= minKey);
    }

    t.textContent = `${hLabel} - 経常利益+減価償却費`;

    if (years.length === 0) {
      const e = document.createElement("div");
      e.className = "dashboard-chart-empty";
      e.textContent = "PLデータなし";
      cdiv.appendChild(e);
      return;
    }

    // 各年度: 経常利益(青系) と 減価償却(黄系) を同じstack内で積み上げ
    // 月次/累積はグローバル dashboardMode に従う
    const total = years.length;
    const datasets = [];
    years.forEach((y, idx) => {
      const yDisp = yearLabelDisplay(y);
      const t = total <= 1 ? 0 : idx / (total - 1);
      const blueColor = `hsl(210, ${70 - t * 35}%, ${25 + t * 63}%)`;
      const yellowColor = `hsl(45, ${85 - t * 40}%, ${45 + t * 40}%)`;

      // 経常利益 月次データ
      const keijoMonths = itemBlocks[0]?.[y] || {};
      const keijoRaw = new Array(12).fill(null);
      for (let m = 0; m < 12; m++) {
        const v = keijoMonths[String(m)];
        if (typeof v === "number") keijoRaw[m] = v;
      }
      // 減価償却 月次 (新年度=実月次 / 旧年度=÷12 のハイブリッド)
      const genkaBlock = itemBlocks[1];
      const genkaRaw = genkaBlock ? depreciationMonthlyArray(genkaBlock, y) : new Array(12).fill(null);

      const keijoData = dashboardMode === "cumulative" ? cumulative(keijoRaw) : keijoRaw;
      const genkaData = dashboardMode === "cumulative" ? cumulative(genkaRaw) : genkaRaw;

      datasets.push({
        label: `${yDisp} 経常利益`,
        data: keijoData,
        backgroundColor: blueColor,
        borderColor: blueColor,
        borderWidth: 1,
        stack: `y${idx}`,
      });
      datasets.push({
        label: `${yDisp} 減価償却`,
        data: genkaData,
        backgroundColor: yellowColor,
        borderColor: yellowColor,
        borderWidth: 1,
        stack: `y${idx}`,
      });
    });
    const monthLabels = getMonthLabels(detectSheetMonthStart(sheet));
    const c = new Chart(canvas.getContext("2d"), {
      type: "bar",
      data: { labels: monthLabels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false, animation: false,
        plugins: {
          legend: {
            position: "bottom",
            labels: {
              boxWidth: 12, font: { size: 11 },
              // 凡例は「経常利益(青)」「減価償却(黄)」の2つだけに集約
              generateLabels: () => [
                { text: "経常利益", fillStyle: "rgba(20,60,120,0.92)", strokeStyle: "rgba(20,60,120,0.92)" },
                { text: "減価償却（新=実月次/旧=÷12）", fillStyle: "rgba(210,170,60,0.9)", strokeStyle: "rgba(210,170,60,0.9)" },
              ],
            },
          },
          tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}` } },
        },
        scales: {
          x: { stacked: true, ticks: { font: { size: 9 } } },
          y: { stacked: true, ticks: { font: { size: 9 }, callback: v => v.toLocaleString("ja-JP") } },
        },
      },
    });
    dashboardCharts.push(c);
  });
  grid.appendChild(sec);
}

// ---------- ⑫ 簡易キャッシュフロー (営業/投資/財務) ----------
// 営業CF = 経常利益 + 減価償却費 (簡易・運転資本変動は無視)
// 投資CF = -(Δ有形固定資産簿価 + 減価償却費) = -CAPEX
// 財務CF = Δ長期借入金 + Δ短期借入金
function renderCFSection(grid, section) {
  const sec = document.createElement("section");
  sec.className = "dashboard-section";
  appendSectionHeader(sec, section);

  // 注記
  const note = document.createElement("p");
  note.style.cssText = "color:#666;font-size:11px;margin:4px 0 8px;";
  note.innerHTML =
    "<b>【間接法CF・計算式】</b> 起点=税引前当期純損益 (なければ経常損益で代替)<br>" +
    "<b>営業CF</b> = 税引前当期純損益 + 減価償却費 + Δ各種引当金(貸倒/賞与/退職給付)<br>" +
    "&nbsp;&nbsp;&nbsp;− 法人税等支払額(=発生額−Δ未払法人税等) − Δ売上債権 − Δ棚卸資産<br>" +
    "&nbsp;&nbsp;&nbsp;+ Δ未払金 ± Δその他運転資本(前払/前受/預り金/未収入金等)<br>" +
    "<b>投資CF</b> = −(Δ有形固定資産簿価 + 減価償却費) − Δ投資その他資産(出資金/敷金/無形等) + 固定資産売却益<br>" +
    "<b>財務CF</b> = Δ長期借入金 + Δ短期借入金 − Δ役員貸付金 − Δ関係会社債権 + Δ関係会社債務<br>" +
    "<span style='color:#999'>※引当金・運転資本変動はPCA仕訳明細(InputSlip)から取得。R7末は決算未確定のため参考値。</span>";
  sec.appendChild(note);

  const chartsDiv = document.createElement("div");
  chartsDiv.className = "dashboard-charts";
  sec.appendChild(chartsDiv);

  // PL: 起点利益 (税引前当期純損益を優先・なければ経常損益) ・減価償却費・当期純損益
  const pretax = extractHoujinAnnualSum(section.houjinPlSheets, ["税引前当期純利益", "税引前当期純損益"]);
  const keijo = extractHoujinAnnualSum(section.houjinPlSheets, ["経常利益", "経常損益"]);
  const netInc = extractHoujinAnnualSum(section.houjinPlSheets, ["当期純利益", "当期純損益"]);
  const genka = extractHoujinAnnualSum(section.houjinPlSheets, ["減価償却費"]);
  // BS: 有形固定資産・借入金・運転資本・グループ間取引
  const bvAsset = extractHoujinYearEnd(section.houjinBsSheets, "有形固定資産（簿価）");
  const longLoan = extractHoujinYearEnd(section.houjinBsSheets, "長期借入金");
  const shortLoan = extractHoujinYearEnd(section.houjinBsSheets, "短期借入金");
  const recv = extractHoujinYearEnd(section.houjinBsSheets, "医業未収金");
  const inv = extractHoujinYearEnd(section.houjinBsSheets, "棚卸資産");
  const payable = extractHoujinYearEnd(section.houjinBsSheets, "未払金・未払費用");
  const offLoan = extractHoujinYearEnd(section.houjinBsSheets, "役員貸付金");
  const grpRecv = extractHoujinYearEnd(section.houjinBsSheets, "関係会社債権");
  const grpPay = extractHoujinYearEnd(section.houjinBsSheets, "関係会社債務");

  const houjinLabels = Object.keys(section.houjinPlSheets);
  houjinLabels.forEach((hLabel, idx) => {
    // X軸はBS年度に限定 (運転資本変動・投資CF・財務CF が計算できる年度のみ意味があるため)
    const hYears = new Set();
    [bvAsset, longLoan, shortLoan, recv, inv, payable, offLoan, grpRecv, grpPay].forEach(d => {
      const v = d.result[hLabel] || {};
      Object.keys(v).forEach(y => hYears.add(y));
    });
    let sortedYears = dedupeYearsByDisplay(Array.from(hYears)).sort((a, b) => yearOrderKey(a) - yearOrderKey(b));
    if (section.minYear) {
      const minKey = yearOrderKey(section.minYear === "R2" ? "2.4-3.3" : section.minYear);
      sortedYears = sortedYears.filter(y => yearOrderKey(y) >= minKey);
    }
    const yearDisplays = sortedYears.map(y => yearLabelDisplay(y));

    const cdiv = document.createElement("div");
    cdiv.className = "dashboard-chart";
    const t = document.createElement("div");
    t.className = "dashboard-chart-title";
    t.textContent = `${hLabel} - 簡易CF（年度概算）`;
    cdiv.appendChild(t);
    const wrap = document.createElement("div");
    wrap.className = "dashboard-chart-canvas";
    const canvas = document.createElement("canvas");
    wrap.appendChild(canvas);
    cdiv.appendChild(wrap);
    chartsDiv.appendChild(cdiv);

    const opCF = [];
    const invCF = [];
    const finCF = [];
    // 両年度ともBSにある時だけ差分を加算するヘルパー
    const diff = (src, y, prevY) => {
      const a = src.result[hLabel]?.[y];
      const b = src.result[hLabel]?.[prevY];
      return (typeof a === "number" && typeof b === "number") ? (a - b) : 0;
    };
    const has = (src, y, prevY) => {
      return typeof src.result[hLabel]?.[y] === "number" && typeof src.result[hLabel]?.[prevY] === "number";
    };
    // 純増減ブロック (InputSlip Dr/Cr由来・年合計は月11に格納) を読むヘルパー
    const bsSheet = state.data.sheets[section.houjinBsSheets[hLabel]];
    const netChg = (blockName, y) => {
      const blk = bsSheet?.blocks?.[blockName];
      const v = blk?.[y]?.["11"];
      return typeof v === "number" ? v : 0;
    };
    sortedYears.forEach((y, yi) => {
      const prevY = yi > 0 ? sortedYears[yi - 1] : null;
      // 起点利益: 税引前優先・なければ経常損益
      const pt = pretax.result[hLabel]?.[y];
      const ke = keijo.result[hLabel]?.[y];
      const baseProfit = typeof pt === "number" ? pt : (typeof ke === "number" ? ke : null);
      const usingPretax = typeof pt === "number";
      const g = genka.result[hLabel]?.[y];
      const ni = netInc.result[hLabel]?.[y];

      // 営業CF (間接法・拡張版)
      if (typeof baseProfit === "number") {
        let op = baseProfit;
        if (typeof g === "number") op += g;
        // 法人税等支払額 ≒ 税引前 − 当期純 − Δ未払法人税等 (発生額→支払額へ補正)
        if (usingPretax && typeof ni === "number") {
          op -= (pt - ni);                       // 発生額を控除
          op += netChg("未払法人税等_純増減", y);  // 未払増=未払い分を戻す
        }
        // 引当金 (非現金費用・足し戻し)
        op += netChg("貸倒引当金_純増減", y);
        op += netChg("賞与引当金_純増減", y);
        op += netChg("退職給付引当金_純増減", y);
        // 主要運転資本 (BS年度末残高差分)
        if (prevY) {
          op -= diff(recv, y, prevY);
          op -= diff(inv, y, prevY);
          op += diff(payable, y, prevY);
        }
        // その他運転資本 (純増減ベース・資産増=CF減/負債増=CF増)
        op -= netChg("その他流動資産_純増減", y);
        op += netChg("その他流動負債_純増減", y);
        // 固定資産売却損益は非営業 → 営業CFから除外 (投資CFへ振替)
        op -= netChg("固定資産売却損益_純増減", y);
        opCF.push(op);
      } else opCF.push(null);

      // 投資CF = -(Δ有形固定資産 + 減価償却費) − Δ投資その他資産 + 固定資産売却益
      if (prevY && has(bvAsset, y, prevY)) {
        const dBV = bvAsset.result[hLabel][y] - bvAsset.result[hLabel][prevY];
        const gv = typeof g === "number" ? g : 0;
        let iv = -(dBV + gv);
        iv -= netChg("投資その他資産_純増減", y);     // 投資資産増=投資CF減
        iv += netChg("固定資産売却損益_純増減", y);   // 売却益=売却収入の上乗せ分
        invCF.push(iv);
      } else invCF.push(null);

      // 財務CF = Δ借入金 − Δ役員貸付金 − Δ関係会社債権 + Δ関係会社債務
      if (prevY) {
        let fin = 0;
        let any = false;
        if (has(longLoan, y, prevY))  { fin += diff(longLoan, y, prevY); any = true; }
        if (has(shortLoan, y, prevY)) { fin += diff(shortLoan, y, prevY); any = true; }
        if (has(offLoan, y, prevY))   { fin -= diff(offLoan, y, prevY); any = true; } // 役員貸付↑=CF↓ (法人から流出)
        if (has(grpRecv, y, prevY))   { fin -= diff(grpRecv, y, prevY); any = true; } // 関係会社債権↑=CF↓
        if (has(grpPay, y, prevY))    { fin += diff(grpPay, y, prevY); any = true; }  // 関係会社債務↑=CF↑
        finCF.push(any ? fin : null);
      } else finCF.push(null);
    });

    const datasets = [
      { label: "営業CF", data: opCF, backgroundColor: "rgba(46,134,193,0.85)", borderColor: "rgba(46,134,193,1)", borderWidth: 1 },
      { label: "投資CF", data: invCF, backgroundColor: "rgba(231,76,60,0.75)", borderColor: "rgba(231,76,60,1)", borderWidth: 1 },
      { label: "財務CF", data: finCF, backgroundColor: "rgba(241,196,15,0.85)", borderColor: "rgba(241,196,15,1)", borderWidth: 1 },
    ];

    if (datasets.every(d => d.data.every(v => v === null))) {
      const e = document.createElement("div");
      e.className = "dashboard-chart-empty";
      e.textContent = "CF算出不可（データ不足）";
      cdiv.appendChild(e);
      return;
    }

    const c = new Chart(canvas.getContext("2d"), {
      type: "bar",
      data: { labels: yearDisplays, datasets },
      options: {
        responsive: true, maintainAspectRatio: false, animation: false,
        plugins: {
          legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}` } },
        },
        scales: {
          x: { ticks: { font: { size: 10 } } },
          y: { ticks: { font: { size: 9 }, callback: v => v.toLocaleString("ja-JP") } },
        },
      },
    });
    dashboardCharts.push(c);
  });
  grid.appendChild(sec);
}

// ---------- ⑬ 借入金増減 (新規借入 vs 返済) ----------
// InputSlip Dr/Cr分離で取得した「短期借入_新規/返済」「長期借入_新規/返済」を集計
// 新規借入(+, 上向き) と 返済(-, 下向き) を年度別に表示
function renderBorrowFlowSection(grid, section) {
  const sec = document.createElement("section");
  sec.className = "dashboard-section";
  appendSectionHeader(sec, section);
  const note = document.createElement("p");
  note.style.cssText = "color:#666;font-size:11px;margin:4px 0 8px;";
  note.textContent = "※PCA仕訳明細(InputSlip)のDr/Cr分離で取得。新規借入(+)=期中の借入実行額、返済(−)=期中の返済額。短期+長期を合算。BS残高差分では見えない『借り換え』も把握できる。";
  sec.appendChild(note);
  const chartsDiv = document.createElement("div");
  chartsDiv.className = "dashboard-charts";
  sec.appendChild(chartsDiv);

  const houjinLabels = Object.keys(section.houjinSheets);
  houjinLabels.forEach((hLabel) => {
    const sheetName = section.houjinSheets[hLabel];
    const sheet = state.data.sheets[sheetName];

    const cdiv = document.createElement("div");
    cdiv.className = "dashboard-chart";
    const t = document.createElement("div");
    t.className = "dashboard-chart-title";
    t.textContent = `${hLabel} - 借入金 新規 vs 返済（年度別）`;
    cdiv.appendChild(t);
    const wrap = document.createElement("div");
    wrap.className = "dashboard-chart-canvas";
    const canvas = document.createElement("canvas");
    wrap.appendChild(canvas);
    cdiv.appendChild(wrap);
    chartsDiv.appendChild(cdiv);

    const blocks = sheet?.blocks || {};
    const newBorrowBlocks = ["短期借入金_新規借入", "長期借入金_新規借入"];
    const repayBlocks = ["短期借入金_返済", "長期借入金_返済"];

    // 年度収集
    const years = new Set();
    [...newBorrowBlocks, ...repayBlocks].forEach(bk => {
      Object.keys(blocks[bk] || {}).forEach(y => years.add(y));
    });
    const sortedYears = dedupeYearsByDisplay(Array.from(years)).sort((a, b) => yearOrderKey(a) - yearOrderKey(b));

    if (sortedYears.length === 0) {
      const e = document.createElement("div");
      e.className = "dashboard-chart-empty";
      e.textContent = "借入金増減データなし";
      cdiv.appendChild(e);
      return;
    }

    const annualOf = (blockName, y) => {
      const blk = blocks[blockName];
      if (!blk || !blk[y]) return 0;
      // 年合計は月11に格納
      const v = blk[y]["11"];
      return typeof v === "number" ? v : 0;
    };

    const newData = sortedYears.map(y => newBorrowBlocks.reduce((s, bk) => s + annualOf(bk, y), 0));
    const repayData = sortedYears.map(y => -repayBlocks.reduce((s, bk) => s + annualOf(bk, y), 0)); // 返済はマイナス表示

    if (newData.every(v => v === 0) && repayData.every(v => v === 0)) {
      const e = document.createElement("div");
      e.className = "dashboard-chart-empty";
      e.textContent = "借入金の動きなし";
      cdiv.appendChild(e);
      return;
    }

    const yearDisplays = sortedYears.map(y => yearLabelDisplay(y));
    const c = new Chart(canvas.getContext("2d"), {
      type: "bar",
      data: {
        labels: yearDisplays,
        datasets: [
          { label: "新規借入", data: newData, backgroundColor: "rgba(46,134,193,0.85)", borderColor: "rgba(46,134,193,1)", borderWidth: 1 },
          { label: "返済", data: repayData, backgroundColor: "rgba(231,76,60,0.8)", borderColor: "rgba(231,76,60,1)", borderWidth: 1 },
        ],
      },
      options: {
        responsive: true, maintainAspectRatio: false, animation: false,
        plugins: {
          legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmt(Math.abs(ctx.parsed.y))}` } },
        },
        scales: {
          x: { ticks: { font: { size: 10 } } },
          y: { ticks: { font: { size: 9 }, callback: v => v.toLocaleString("ja-JP") } },
        },
      },
    });
    dashboardCharts.push(c);
  });
  grid.appendChild(sec);
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

// データ正規化：ブロック内に紛れ込んだ非年度キー（"_"始まり: __meta__ 等の出所メタ）を除去。
// 年度列挙（getMatrix等）は block のキーを全て年度として扱うため、ブロック直下のメタが
// 「__meta__」という空の年度系列として凡例・表に出てしまうのを防ぐ防御線（データ側の根治と二重化）。
function sanitizeData(data) {
  if (!data || !data.sheets) return;
  for (const sheet of Object.values(data.sheets)) {
    const blocks = sheet && sheet.blocks;
    if (!blocks) continue;
    for (const block of Object.values(blocks)) {
      if (block && typeof block === "object" && !Array.isArray(block)) {
        for (const k of Object.keys(block)) {
          if (k.startsWith("_")) delete block[k];
        }
      }
    }
  }
}

// ---------- 起動 ----------
async function init() {
  // データ読み込み
  try {
    const res = await fetch(DATA_URL);
    state.data = await res.json();
    sanitizeData(state.data);
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
