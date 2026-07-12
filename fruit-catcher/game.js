/* 體感接水果 —— 前鏡頭手部偵測小遊戲
 *
 * 手部偵測優先使用 MediaPipe HandLandmarker（CDN 載入）；
 * 若載入失敗（離線等），自動退回「畫面移動偵測」模式，
 * 以移動像素的重心當作手的位置，遊戲照樣能玩。
 */

const GAME_SECONDS = 60;
const HAND_RADIUS = 44; // 手掌判定半徑（px，會依螢幕縮放）

// 水果表：越大顆分數越高
const FRUITS = [
  { emoji: "🍉", name: "西瓜", r: 46, points: 50, color: "#f87171" },
  { emoji: "🍍", name: "鳳梨", r: 40, points: 40, color: "#fbbf24" },
  { emoji: "🍎", name: "蘋果", r: 32, points: 30, color: "#ef4444" },
  { emoji: "🍌", name: "香蕉", r: 30, points: 25, color: "#fde047" },
  { emoji: "🍊", name: "橘子", r: 28, points: 25, color: "#fb923c" },
  { emoji: "🍇", name: "葡萄", r: 26, points: 20, color: "#a78bfa" },
  { emoji: "🥝", name: "奇異果", r: 24, points: 20, color: "#84cc16" },
  { emoji: "🍓", name: "草莓", r: 20, points: 15, color: "#fb7185" },
];
const BOMB = { emoji: "💣", name: "炸彈", r: 30, points: -40, color: "#64748b", hazard: true };
const DURIAN = { name: "榴槤", r: 38, points: -30, color: "#a3e635", hazard: true, durian: true };

// ---------- DOM ----------
const video = document.getElementById("video");
const canvas = document.getElementById("game");
const ctx = canvas.getContext("2d");
const hud = document.getElementById("hud");
const hudScore = document.querySelector("#hudScore b");
const hudTime = document.querySelector("#hudTime b");
const hudCombo = document.getElementById("hudCombo");
const startOverlay = document.getElementById("startOverlay");
const endOverlay = document.getElementById("endOverlay");
const statusEl = document.getElementById("status");
const legendEl = document.getElementById("legend");

// 開始畫面的分數對照表
legendEl.innerHTML =
  FRUITS.map((f) => `<span>${f.emoji} ${f.points}</span>`).join("") +
  `<span>💣 ${BOMB.points}</span><span>🟢榴槤 ${DURIAN.points}</span>`;

// ---------- 狀態 ----------
let landmarker = null; // MediaPipe HandLandmarker
let useMotionFallback = false;
let running = false;
let rafId = 0;

let score = 0;
let combo = 0;
let lastCatchAt = 0;
let timeLeft = GAME_SECONDS;
let lastTs = 0;
let spawnTimer = 0;
let elapsed = 0;
let flash = 0; // 踩到炸彈的紅色閃屏
let shake = 0;

let entities = []; // 飛行中的水果/炸彈
let particles = [];
let floaters = []; // 漂浮的加減分文字
let hands = []; // [{x, y}] 螢幕座標

function resize() {
  canvas.width = window.innerWidth * devicePixelRatio;
  canvas.height = window.innerHeight * devicePixelRatio;
  ctx.setTransform(devicePixelRatio, 0, 0, devicePixelRatio, 0, 0);
}
window.addEventListener("resize", resize);
resize();

const W = () => window.innerWidth;
const H = () => window.innerHeight;
// 依螢幕大小縮放（以 420px 寬為基準的手機畫面）
const S = () => Math.max(0.7, Math.min(W(), H()) / 420);

// ---------- 音效（WebAudio 合成，無外部檔案） ----------
let audioCtx = null;
function beep(freq, dur = 0.12, type = "sine", gain = 0.15) {
  try {
    audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
    const o = audioCtx.createOscillator();
    const g = audioCtx.createGain();
    o.type = type;
    o.frequency.value = freq;
    g.gain.setValueAtTime(gain, audioCtx.currentTime);
    g.gain.exponentialRampToValueAtTime(0.001, audioCtx.currentTime + dur);
    o.connect(g).connect(audioCtx.destination);
    o.start();
    o.stop(audioCtx.currentTime + dur);
  } catch (_) { /* 無聲也能玩 */ }
}
const sfxCatch = (pts) => beep(500 + pts * 8, 0.12, "triangle", 0.18);
const sfxBad = () => { beep(120, 0.3, "sawtooth", 0.22); beep(80, 0.4, "square", 0.15); };

// ---------- 相機 ----------
async function setupCamera() {
  const stream = await navigator.mediaDevices.getUserMedia({
    video: { facingMode: "user", width: { ideal: 960 }, height: { ideal: 720 } },
    audio: false,
  });
  video.srcObject = stream;
  await new Promise((res) => (video.onloadedmetadata = res));
  await video.play();
}

// ---------- 手部偵測（MediaPipe） ----------
async function setupHandLandmarker() {
  const vision = await import(
    "https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.14"
  );
  const fileset = await vision.FilesetResolver.forVisionTasks(
    "https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.14/wasm"
  );
  landmarker = await vision.HandLandmarker.createFromOptions(fileset, {
    baseOptions: {
      modelAssetPath:
        "https://storage.googleapis.com/mediapipe-models/hand_landmarker/hand_landmarker/float16/1/hand_landmarker.task",
      delegate: "GPU",
    },
    numHands: 2,
    runningMode: "VIDEO",
  });
}

// 把影像座標（0~1）換算成螢幕座標（考慮 object-fit: cover 的裁切 + 鏡像）
function videoToScreen(nx, ny) {
  const vw = video.videoWidth, vh = video.videoHeight;
  const scale = Math.max(W() / vw, H() / vh);
  const dw = vw * scale, dh = vh * scale;
  const ox = (W() - dw) / 2, oy = (H() - dh) / 2;
  return { x: W() - (nx * dw + ox), y: ny * dh + oy }; // x 鏡像
}

function detectHands(ts) {
  hands = [];
  if (landmarker) {
    const result = landmarker.detectForVideo(video, ts);
    for (const lm of result.landmarks || []) {
      // 手掌中心：手腕(0) 與中指根部(9) 的中點
      const cx = (lm[0].x + lm[9].x) / 2;
      const cy = (lm[0].y + lm[9].y) / 2;
      hands.push(videoToScreen(cx, cy));
    }
  } else if (useMotionFallback) {
    const p = detectMotion();
    if (p) hands.push(p);
  }
}

// ---------- 備援：移動偵測（畫格差分的重心） ----------
const motionCanvas = document.createElement("canvas");
motionCanvas.width = 80;
motionCanvas.height = 60;
const mctx = motionCanvas.getContext("2d", { willReadFrequently: true });
let prevFrame = null;
let smoothedHand = null;

function detectMotion() {
  const mw = motionCanvas.width, mh = motionCanvas.height;
  mctx.drawImage(video, 0, 0, mw, mh);
  const frame = mctx.getImageData(0, 0, mw, mh);
  if (!prevFrame) { prevFrame = frame; return null; }

  let sumX = 0, sumY = 0, count = 0;
  const a = frame.data, b = prevFrame.data;
  for (let i = 0; i < a.length; i += 4) {
    const diff = Math.abs(a[i] - b[i]) + Math.abs(a[i + 1] - b[i + 1]) + Math.abs(a[i + 2] - b[i + 2]);
    if (diff > 90) {
      const px = (i / 4) % mw, py = ((i / 4) / mw) | 0;
      sumX += px; sumY += py; count++;
    }
  }
  prevFrame = frame;
  if (count < 12) return smoothedHand; // 沒明顯移動就停在原地

  const target = videoToScreen(sumX / count / mw, sumY / count / mh);
  smoothedHand = smoothedHand
    ? { x: smoothedHand.x * 0.6 + target.x * 0.4, y: smoothedHand.y * 0.6 + target.y * 0.4 }
    : target;
  return smoothedHand;
}

// ---------- 生成水果 / 炸彈 ----------
function spawnEntity() {
  const roll = Math.random();
  let type;
  if (roll < 0.12) type = BOMB;
  else if (roll < 0.22) type = DURIAN;
  else type = FRUITS[(Math.random() * FRUITS.length) | 0];

  const fromLeft = Math.random() < 0.5;
  const s = S();
  const r = type.r * s;
  const y = H() * (0.25 + Math.random() * 0.55);
  const speed = (W() / 3.2) * (1 + elapsed / 90) * (0.85 + Math.random() * 0.4);

  entities.push({
    type,
    r,
    x: fromLeft ? -r : W() + r,
    y,
    vx: fromLeft ? speed : -speed,
    vy: -(60 + Math.random() * 140) * s, // 先微微上拋
    g: (140 + Math.random() * 80) * s,   // 再受重力落下 → 弧線
    rot: Math.random() * Math.PI * 2,
    vr: (Math.random() - 0.5) * 3,
    caught: false,
    fade: 1,
  });
}

// ---------- 特效 ----------
function burst(x, y, color, n = 14) {
  for (let i = 0; i < n; i++) {
    const ang = Math.random() * Math.PI * 2;
    const sp = 80 + Math.random() * 260;
    particles.push({
      x, y,
      vx: Math.cos(ang) * sp,
      vy: Math.sin(ang) * sp - 60,
      life: 0.6 + Math.random() * 0.3,
      t: 0,
      color,
      r: 2 + Math.random() * 4,
    });
  }
}

function addFloater(x, y, text, color) {
  floaters.push({ x, y, text, color, t: 0 });
}

// ---------- 碰撞與計分 ----------
function handleCatches(now) {
  for (const e of entities) {
    if (e.caught) continue;
    for (const h of hands) {
      const dx = e.x - h.x, dy = e.y - h.y;
      if (Math.hypot(dx, dy) < e.r * 0.95 + HAND_RADIUS * S()) {
        e.caught = true;
        if (e.type.hazard) {
          score = Math.max(0, score + e.type.points);
          combo = 0;
          flash = 1;
          shake = 1;
          sfxBad();
          addFloater(e.x, e.y, `${e.type.points} ${e.type.name}！`, "#f87171");
          burst(e.x, e.y, e.type.color, 22);
          if (navigator.vibrate) navigator.vibrate(200);
        } else {
          combo = now - lastCatchAt < 2000 ? combo + 1 : 1;
          lastCatchAt = now;
          const bonus = Math.min(combo - 1, 5) * 5;
          const gained = e.type.points + bonus;
          score += gained;
          sfxCatch(e.type.points);
          addFloater(e.x, e.y, `+${gained}`, "#4ade80");
          burst(e.x, e.y, e.type.color);
          if (navigator.vibrate) navigator.vibrate(30);
        }
        break;
      }
    }
  }
  hudScore.textContent = score;
  hudCombo.textContent = combo >= 2 ? `🔥 連擊 x${combo}` : "";
}

// ---------- 畫榴槤（沒有 emoji，自己畫一顆刺刺的） ----------
function drawDurian(x, y, r, rot) {
  ctx.save();
  ctx.translate(x, y);
  ctx.rotate(rot * 0.3);
  const spikes = 16;
  ctx.beginPath();
  for (let i = 0; i <= spikes * 2; i++) {
    const ang = (i / (spikes * 2)) * Math.PI * 2;
    const rad = i % 2 === 0 ? r : r * 0.78;
    ctx.lineTo(Math.cos(ang) * rad, Math.sin(ang) * rad);
  }
  ctx.closePath();
  const grad = ctx.createRadialGradient(0, -r * 0.3, r * 0.2, 0, 0, r);
  grad.addColorStop(0, "#d9f99d");
  grad.addColorStop(1, "#65a30d");
  ctx.fillStyle = grad;
  ctx.fill();
  ctx.strokeStyle = "#3f6212";
  ctx.lineWidth = 2;
  ctx.stroke();
  // 蒂頭
  ctx.fillStyle = "#78350f";
  ctx.fillRect(-3, -r - 8, 6, 10);
  ctx.restore();
}

// ---------- 主迴圈 ----------
function loop(ts) {
  if (!running) return;
  rafId = requestAnimationFrame(loop);

  const dt = Math.min(0.05, lastTs ? (ts - lastTs) / 1000 : 0.016);
  lastTs = ts;
  elapsed += dt;

  // 倒數
  timeLeft -= dt;
  hudTime.textContent = Math.max(0, Math.ceil(timeLeft));
  if (timeLeft <= 0) { endGame(); return; }

  detectHands(ts);

  // 生成節奏：由 1.1 秒一顆漸進加快到 0.45 秒一顆
  spawnTimer -= dt;
  if (spawnTimer <= 0) {
    spawnEntity();
    spawnTimer = Math.max(0.45, 1.1 - elapsed * 0.011);
  }

  // 物理更新
  for (const e of entities) {
    if (e.caught) { e.fade -= dt * 4; continue; }
    e.vy += e.g * dt;
    e.x += e.vx * dt;
    e.y += e.vy * dt;
    e.rot += e.vr * dt;
  }
  entities = entities.filter(
    (e) => e.fade > 0 && e.x > -200 && e.x < W() + 200 && e.y < H() + 200
  );

  handleCatches(ts);

  // ---- 繪製 ----
  ctx.clearRect(0, 0, W(), H());
  ctx.save();
  if (shake > 0) {
    shake = Math.max(0, shake - dt * 3);
    ctx.translate((Math.random() - 0.5) * 14 * shake, (Math.random() - 0.5) * 14 * shake);
  }

  // 水果 / 炸彈 / 榴槤
  for (const e of entities) {
    ctx.save();
    ctx.globalAlpha = Math.max(0, e.fade);
    const scale = e.caught ? 1 + (1 - e.fade) * 0.6 : 1;
    if (e.type.durian) {
      drawDurian(e.x, e.y, e.r * scale, e.rot);
    } else {
      ctx.translate(e.x, e.y);
      ctx.rotate(Math.sin(e.rot) * 0.35);
      ctx.font = `${e.r * 2 * scale}px serif`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(e.type.emoji, 0, 0);
    }
    ctx.restore();
  }

  // 手掌游標
  const hr = HAND_RADIUS * S();
  for (const h of hands) {
    ctx.save();
    ctx.beginPath();
    ctx.arc(h.x, h.y, hr, 0, Math.PI * 2);
    ctx.fillStyle = "rgba(74, 222, 128, 0.18)";
    ctx.fill();
    ctx.lineWidth = 3;
    ctx.strokeStyle = "rgba(74, 222, 128, 0.9)";
    ctx.shadowColor = "#4ade80";
    ctx.shadowBlur = 18;
    ctx.stroke();
    ctx.shadowBlur = 0;
    ctx.font = `${hr}px serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("🖐️", h.x, h.y);
    ctx.restore();
  }

  // 粒子
  for (const p of particles) {
    p.t += dt;
    p.x += p.vx * dt;
    p.y += p.vy * dt;
    p.vy += 500 * dt;
    ctx.globalAlpha = Math.max(0, 1 - p.t / p.life);
    ctx.fillStyle = p.color;
    ctx.beginPath();
    ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
    ctx.fill();
  }
  particles = particles.filter((p) => p.t < p.life);
  ctx.globalAlpha = 1;

  // 加減分文字
  for (const f of floaters) {
    f.t += dt;
    ctx.globalAlpha = Math.max(0, 1 - f.t / 1);
    ctx.font = `bold ${26 * S()}px sans-serif`;
    ctx.textAlign = "center";
    ctx.fillStyle = f.color;
    ctx.strokeStyle = "rgba(0,0,0,0.6)";
    ctx.lineWidth = 4;
    const fy = f.y - f.t * 70;
    ctx.strokeText(f.text, f.x, fy);
    ctx.fillText(f.text, f.x, fy);
  }
  floaters = floaters.filter((f) => f.t < 1);
  ctx.globalAlpha = 1;
  ctx.restore();

  // 炸彈紅色閃屏
  if (flash > 0) {
    flash = Math.max(0, flash - dt * 2.5);
    ctx.fillStyle = `rgba(239, 68, 68, ${flash * 0.35})`;
    ctx.fillRect(0, 0, W(), H());
  }
}

// ---------- 遊戲流程 ----------
function resetGame() {
  score = 0;
  combo = 0;
  timeLeft = GAME_SECONDS;
  elapsed = 0;
  lastTs = 0;
  spawnTimer = 0.4;
  entities = [];
  particles = [];
  floaters = [];
  flash = 0;
  shake = 0;
  prevFrame = null;
  smoothedHand = null;
  hudScore.textContent = "0";
  hudTime.textContent = GAME_SECONDS;
  hudCombo.textContent = "";
}

function endGame() {
  running = false;
  cancelAnimationFrame(rafId);
  hud.classList.add("hidden");

  const best = Math.max(score, Number(localStorage.getItem("fruitCatcherBest") || 0));
  localStorage.setItem("fruitCatcherBest", best);
  document.getElementById("finalScore").textContent = score;
  document.getElementById("bestScore").textContent = best;
  document.getElementById("endComment").textContent =
    score >= 600 ? "🏆 水果大師！反應快到不可思議！" :
    score >= 350 ? "💪 很厲害！再快一點點就是大師了！" :
    score >= 150 ? "🙂 不錯喔，多練習閃開炸彈和榴槤！" :
    "😅 慢慢來，記得用手掌去碰水果～";
  endOverlay.classList.remove("hidden");
}

function startRound() {
  resetGame();
  startOverlay.classList.add("hidden");
  endOverlay.classList.add("hidden");
  hud.classList.remove("hidden");
  running = true;
  rafId = requestAnimationFrame(loop);
}

async function init() {
  const btn = document.getElementById("btnStart");
  btn.disabled = true;
  try {
    statusEl.textContent = "📷 開啟相機中…請允許使用相機";
    await setupCamera();
  } catch (err) {
    statusEl.textContent = "⚠️ 無法開啟相機：請確認已允許相機權限，並使用 https 開啟本頁。";
    btn.disabled = false;
    return;
  }

  try {
    statusEl.textContent = "🖐️ 載入手部偵測模型中…（第一次會多花幾秒）";
    await setupHandLandmarker();
    statusEl.textContent = "";
  } catch (err) {
    // 離線或 CDN 被擋 → 改用移動偵測
    useMotionFallback = true;
    statusEl.textContent = "";
    console.warn("HandLandmarker 載入失敗，改用移動偵測模式", err);
  }

  startRound();
}

document.getElementById("btnStart").addEventListener("click", init);
document.getElementById("btnRestart").addEventListener("click", startRound);
