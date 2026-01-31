// =============================
// Config
// =============================
const API_BASE = "https://YOUR_WORKER_SUBDOMAIN.workers.dev"; // <-- change me

// =============================
// Quiz Data
// =============================
const QUESTIONS = [
  {
    title: "Q1｜市場大跌 20%，你的第一反應是？",
    options: [
      { key: "A", text: "先停手，重新檢查整體配置" },
      { key: "B", text: "機會來了，分批慢慢加" },
      { key: "C", text: "All in，這波不拼不行" },
      { key: "D", text: "觀望，等趨勢明朗再說" }
    ]
  },
  {
    title: "Q2｜你最能接受哪種投資節奏？",
    options: [
      { key: "A", text: "慢慢來，十年不嫌久" },
      { key: "B", text: "有紀律的進出場" },
      { key: "C", text: "快、狠、準，錯了就換" },
      { key: "D", text: "不急，等對的時機" }
    ]
  },
  {
    title: "Q3｜你最怕的是？",
    options: [
      { key: "A", text: "資產大幅回撤" },
      { key: "B", text: "買太少、錯過行情" },
      { key: "C", text: "錯過翻倍機會" },
      { key: "D", text: "進錯場、站錯邊" }
    ]
  },
  {
    title: "Q4｜如果只能選一種策略？",
    options: [
      { key: "A", text: "資產配置＋長期持有" },
      { key: "B", text: "趨勢投資＋風控" },
      { key: "C", text: "波段／事件型操作" },
      { key: "D", text: "價值低估後耐心等待" }
    ]
  },
  {
    title: "Q5｜朋友問你投資建議時，你通常會說？",
    options: [
      { key: "A", text: "先顧好風險，不要想賺快錢" },
      { key: "B", text: "照紀律做，比判斷重要" },
      { key: "C", text: "這波不敢下就別玩了" },
      { key: "D", text: "時機不對，先等等" }
    ]
  },
  {
    title: "Q6｜你對投資的終極目標是？",
    options: [
      { key: "A", text: "長期穩定、睡得著" },
      { key: "B", text: "穩定成長、可複製" },
      { key: "C", text: "快速翻身、改變階級" },
      { key: "D", text: "低風險抓到大機會" }
    ]
  }
];

const RESULT_MAP = {
  A: {
    title: "A 型｜張無忌型投資人（穩、厚、長期內力）",
    desc:
      "你重視『活得久』勝過『賺得快』。你會先顧風險、顧配置，再談報酬。市場越亂，你越能守住節奏。\n\n" +
      "優勢：抗波動、少犯大錯、長期勝率高。\n" +
      "盲點：可能太晚出手、容易錯過爆發段。\n" +
      "建議：用清楚的再平衡規則，讓你在恐慌時也能『有條件加碼』。"
  },
  B: {
    title: "B 型｜令狐沖型投資人（趨勢、紀律、出手俐落）",
    desc:
      "你相信『紀律』比『預測』重要。你願意跟著趨勢走，也能在錯的時候停手。\n\n" +
      "優勢：順勢吃得到大段行情，風控觀念較完整。\n" +
      "盲點：盤整行情容易沒事做或手癢。\n" +
      "建議：設計「盤整期規則」（例如縮倉、觀察清單、分批試單），避免亂出手。"
  },
  C: {
    title: "C 型｜楊過型投資人（重壓、翻身、情緒放大）",
    desc:
      "你敢賭，也敢扛。你追求的是『一次打穿』的效率：對了很爽，錯了很痛。\n\n" +
      "優勢：掌握事件／波段時可能跑很快。\n" +
      "盲點：回撤管理與情緒控管是生死線。\n" +
      "建議：把「停損/停利/最大回撤」寫成硬規則，避免單次重傷。"
  },
  D: {
    title: "D 型｜郭靖型投資人（慢、忍、厚積薄發）",
    desc:
      "你不搶快，你等『價位對』、『時機對』再出手。你看似慢，其實最難被淘汰。\n\n" +
      "優勢：穩健、耐心、低頻但勝率可觀。\n" +
      "盲點：行情起飛時可能反應太慢。\n" +
      "建議：建立『分批進場框架』，價格到區間就自動執行，不靠臨場情緒。"
  }
};

// =============================
// Rendering
// =============================
const quizEl = document.getElementById("quiz");
const resultEl = document.getElementById("result");
const resultTitleEl = document.getElementById("resultTitle");
const resultDescEl = document.getElementById("resultDesc");
const btnRestart = document.getElementById("btnRestart");
const btnSubmit = document.getElementById("btnSubmit");

const statsEl = document.getElementById("stats");
const statsSummaryEl = document.getElementById("statsSummary");
const barsEl = document.getElementById("bars");
const updatedAtEl = document.getElementById("updatedAt");

let answers = new Array(QUESTIONS.length).fill(null);
let finalKey = null;

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function renderQuiz() {
  quizEl.innerHTML = "";

  QUESTIONS.forEach((q, idx) => {
    const wrap = document.createElement("div");
    wrap.className = "q";

    const h = document.createElement("h3");
    h.textContent = q.title;
    wrap.appendChild(h);

    const opts = document.createElement("div");
    opts.className = "options";

    q.options.forEach((opt) => {
      const label = document.createElement("label");
      label.className = "opt";

      const input = document.createElement("input");
      input.type = "radio";
      input.name = `q${idx}`;
      input.value = opt.key;
      input.checked = answers[idx] === opt.key;
      input.addEventListener("change", () => {
        answers[idx] = opt.key;
        checkDone();
      });

      const span = document.createElement("span");
      span.textContent = opt.text;

      label.appendChild(input);
      label.appendChild(span);
      opts.appendChild(label);
    });

    wrap.appendChild(opts);
    quizEl.appendChild(wrap);
  });

  checkDone();
}

function checkDone() {
  const done = answers.every(Boolean);
  if (done) {
    showResult();
  } else {
    resultEl.classList.add("hidden");
  }
}

function computeFinalKey() {
  const counts = { A: 0, B: 0, C: 0, D: 0 };
  answers.forEach((k) => { if (counts[k] != null) counts[k]++; });

  // tie-breaker priority: A > B > D > C (you can change)
  const order = ["A", "B", "D", "C"];
  let best = "A";
  for (const k of order) {
    if (counts[k] > counts[best]) best = k;
  }
  return best;
}

function showResult() {
  finalKey = computeFinalKey();
  const r = RESULT_MAP[finalKey];

  resultTitleEl.textContent = r.title;
  resultDescEl.textContent = r.desc;

  resultEl.classList.remove("hidden");
}

function resetAll() {
  answers = new Array(QUESTIONS.length).fill(null);
  finalKey = null;
  statsEl.classList.add("hidden");
  btnSubmit.disabled = false;
  renderQuiz();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

// =============================
// API Calls (Anonymous)
// =============================
async function submitAndFetchStats() {
  if (!finalKey) return;

  btnSubmit.disabled = true;

  // Submit result (only finalKey)
  await fetch(`${API_BASE}/submit`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ result: finalKey })
  });

  // Fetch stats
  const res = await fetch(`${API_BASE}/stats`, { method: "GET" });
  const data = await res.json();

  renderStats(data);
}

function renderStats(data) {
  // Expected:
  // { total: number, counts: {A,B,C,D}, updatedAt: ISO }
  const total = Number(data.total || 0);
  const counts = data.counts || { A: 0, B: 0, C: 0, D: 0 };
  const updatedAt = data.updatedAt ? String(data.updatedAt) : null;

  statsSummaryEl.textContent = `目前參與人數：${total} 人`;

  const items = ["A", "B", "C", "D"].map((k) => {
    const n = Number(counts[k] || 0);
    const pct = total > 0 ? Math.round((n / total) * 100) : 0;
    const label = RESULT_MAP[k].title.split("｜")[1]?.split("（")[0] || k;
    return { k, label, n, pct };
  });

  // Sort by pct desc
  items.sort((x, y) => y.pct - x.pct);

  barsEl.innerHTML = items.map((it) => {
    const safeLabel = escapeHtml(it.label);
    return `
      <div class="barRow">
        <div>${safeLabel}</div>
        <div class="bar"><div class="barFill" style="width:${it.pct}%"></div></div>
        <div>${it.pct}%</div>
      </div>
    `;
  }).join("");

  if (updatedAt) {
    updatedAtEl.textContent = `更新時間：${updatedAt}`;
  } else {
    updatedAtEl.textContent = "";
  }

  statsEl.classList.remove("hidden");
}

// =============================
// Events
// =============================
btnRestart.addEventListener("click", resetAll);
btnSubmit.addEventListener("click", submitAndFetchStats);

// init
renderQuiz();
