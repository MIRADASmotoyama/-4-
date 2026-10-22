// フラットな説明用イラスト（SVG → PNG）。画像内には文字を一切描かない。
// 実在企業の設備を描いたものではなく、一般的なイメージ図。
const NAVY = "#12243A", NAVY2 = "#1C3A57", NAVY3 = "#2F5D7C", TEAL = "#067F82",
  CYAN = "#20A8BC", PALE = "#EDF5F7", ORANGE = "#EEA348", STEEL = "#8E9BA6";

// 決定的な疑似乱数（再生成しても同じ絵になる）
function rng(seed) {
  let s = seed;
  return () => ((s = (s * 9301 + 49297) % 233280) / 233280);
}

function wrap(w, h, body, defs = "") {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><defs>${defs}</defs>${body}</svg>`;
}

const skyDefs = (id, a, b) =>
  `<linearGradient id="${id}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${a}"/><stop offset="1" stop-color="${b}"/></linearGradient>`;

// 都市と吊り橋
function cityBridge() {
  const W = 1600, H = 1000, r = rng(7);
  let b = `<rect width="${W}" height="${H}" fill="url(#sky)"/>`;
  b += `<circle cx="1240" cy="230" r="120" fill="${ORANGE}" opacity="0.35"/>`;
  // 奥のビル群
  let x = 0;
  while (x < W) {
    const w = 60 + r() * 90, h = 220 + r() * 330;
    b += `<rect x="${x}" y="${640 - h}" width="${w}" height="${h}" fill="${NAVY3}" opacity="0.55"/>`;
    x += w + 6;
  }
  // 手前のビル群（窓付き）
  x = 20;
  while (x < W) {
    const w = 80 + r() * 110, h = 160 + r() * 300, top = 660 - h;
    b += `<rect x="${x}" y="${top}" width="${w}" height="${h}" fill="${NAVY2}"/>`;
    for (let wy = top + 18; wy < 640; wy += 34)
      for (let wx = x + 12; wx < x + w - 18; wx += 26)
        if (r() > 0.35) b += `<rect x="${wx}" y="${wy}" width="12" height="16" fill="${r() > 0.8 ? ORANGE : "#9FCBD3"}" opacity="0.8"/>`;
    x += w + 14;
  }
  // 水面
  b += `<rect x="0" y="650" width="${W}" height="${H - 650}" fill="#7FB9C4"/>`;
  for (let i = 0; i < 18; i++) b += `<rect x="${r() * W}" y="${690 + r() * 290}" width="${80 + r() * 160}" height="5" rx="2" fill="#FFFFFF" opacity="0.35"/>`;
  // 吊り橋
  const deckY = 700;
  [380, 1220].forEach((tx) => {
    b += `<rect x="${tx - 22}" y="300" width="44" height="${deckY - 300 + 160}" fill="${NAVY}"/>`;
    b += `<rect x="${tx - 34}" y="420" width="68" height="16" fill="${NAVY}"/>`;
    b += `<rect x="${tx - 34}" y="560" width="68" height="16" fill="${NAVY}"/>`;
  });
  b += `<path d="M -40 ${deckY - 10} Q 180 560 380 305 Q 800 690 1220 305 Q 1420 560 1640 ${deckY - 10}" fill="none" stroke="${CYAN}" stroke-width="9"/>`;
  for (let hx = 420; hx < 1200; hx += 40) {
    const t = (hx - 380) / 840, cy = 305 + (1 - Math.pow(2 * t - 1, 2)) * 192;
    b += `<line x1="${hx}" y1="${cy}" x2="${hx}" y2="${deckY}" stroke="${CYAN}" stroke-width="3" opacity="0.8"/>`;
  }
  b += `<rect x="0" y="${deckY}" width="${W}" height="34" fill="${TEAL}"/>`;
  b += `<rect x="0" y="${deckY + 34}" width="${W}" height="10" fill="${NAVY}"/>`;
  [160, 520, 700, 1010, 1380].forEach((cx, i) => {
    b += `<rect x="${cx}" y="${deckY - 22}" width="54" height="22" rx="8" fill="${i % 2 ? ORANGE : "#FFFFFF"}"/>`;
  });
  return wrap(W, H, b, skyDefs("sky", "#CFE6EC", "#F4FAFB"));
}

// 製鉄所（高炉・煙突・溶けた鉄）
function steelMill() {
  const W = 1600, H = 1000;
  let b = `<rect width="${W}" height="${H}" fill="url(#sky)"/>`;
  b += `<rect x="0" y="800" width="${W}" height="200" fill="${NAVY}"/>`;
  // 煙突と蒸気
  [[1080, 180], [1180, 230], [1290, 260]].forEach(([x, top]) => {
    b += `<rect x="${x}" y="${top}" width="46" height="${800 - top}" fill="${NAVY3}"/>`;
    b += `<rect x="${x}" y="${top + 40}" width="46" height="14" fill="${ORANGE}" opacity="0.7"/>`;
    b += `<ellipse cx="${x + 40}" cy="${top - 40}" rx="70" ry="38" fill="#FFFFFF" opacity="0.75"/>`;
    b += `<ellipse cx="${x + 110}" cy="${top - 85}" rx="95" ry="46" fill="#FFFFFF" opacity="0.5"/>`;
  });
  // 工場棟
  b += `<path d="M 60 800 L 60 520 L 300 440 L 300 520 L 540 440 L 540 520 L 780 440 L 780 800 Z" fill="${NAVY2}"/>`;
  for (let i = 0; i < 6; i++) b += `<rect x="${100 + i * 110}" y="600" width="60" height="90" fill="#9FCBD3" opacity="0.5"/>`;
  // 高炉
  b += `<path d="M 880 800 L 900 420 L 940 330 L 1020 330 L 1060 420 L 1080 800 Z" fill="${NAVY}"/>`;
  b += `<rect x="920" y="250" width="120" height="80" fill="${NAVY3}"/>`;
  b += `<rect x="955" y="150" width="50" height="100" fill="${NAVY3}"/>`;
  b += `<path d="M 1040 290 L 1150 200 L 1170 220 L 1060 310 Z" fill="${TEAL}"/>`;
  b += `<line x1="760" y1="560" x2="930" y2="330" stroke="${TEAL}" stroke-width="22"/>`;
  b += `<circle cx="980" cy="700" r="40" fill="${ORANGE}"/><circle cx="980" cy="700" r="70" fill="${ORANGE}" opacity="0.25"/>`;
  // 溶けた鉄の流れと取鍋
  b += `<path d="M 1010 720 Q 1120 760 1180 820" stroke="${ORANGE}" stroke-width="18" fill="none"/>`;
  b += `<path d="M 1130 820 L 1380 820 L 1350 960 L 1160 960 Z" fill="${NAVY3}"/>`;
  b += `<ellipse cx="1255" cy="822" rx="125" ry="22" fill="${ORANGE}"/>`;
  b += `<ellipse cx="1255" cy="800" rx="220" ry="60" fill="${ORANGE}" opacity="0.18"/>`;
  // コイル
  b += `<ellipse cx="420" cy="905" rx="130" ry="70" fill="${STEEL}"/><rect x="290" y="835" width="200" height="140" fill="${STEEL}"/>`;
  b += `<ellipse cx="490" cy="905" rx="40" ry="70" fill="#C9D3DA"/><ellipse cx="490" cy="905" rx="14" ry="26" fill="${NAVY}"/>`;
  return wrap(W, H, b, skyDefs("sky", "#D9E9EE", "#F5F9FA"));
}

// 電子基板（スマホ・半導体）
function chip() {
  const W = 1200, H = 1000, r = rng(3);
  let b = `<rect width="${W}" height="${H}" fill="${NAVY}"/>`;
  for (let i = 0; i < 26; i++) {
    const y = 60 + i * 36, x1 = r() * 300, x2 = 900 + r() * 300;
    b += `<path d="M ${x1} ${y} L ${380} ${y} L ${430} ${y + (r() > 0.5 ? 40 : -40)}" stroke="${ORANGE}" stroke-width="5" fill="none" opacity="0.75"/>`;
    b += `<circle cx="${x1}" cy="${y}" r="9" fill="${ORANGE}" opacity="0.9"/>`;
    b += `<path d="M ${x2} ${y} L ${820} ${y} L ${770} ${y + (r() > 0.5 ? 40 : -40)}" stroke="${CYAN}" stroke-width="5" fill="none" opacity="0.7"/>`;
    b += `<circle cx="${x2}" cy="${y}" r="9" fill="${CYAN}"/>`;
  }
  b += `<rect x="380" y="280" width="440" height="440" rx="24" fill="${NAVY3}"/>`;
  for (let i = 0; i < 10; i++) {
    const p = 310 + i * 40;
    b += `<rect x="${p}" y="250" width="16" height="30" fill="#C9D3DA"/><rect x="${p}" y="720" width="16" height="30" fill="#C9D3DA"/>`;
    b += `<rect x="350" y="${p}" width="30" height="16" fill="#C9D3DA"/><rect x="820" y="${p}" width="30" height="16" fill="#C9D3DA"/>`;
  }
  b += `<rect x="450" y="350" width="300" height="300" rx="14" fill="${NAVY2}" stroke="${CYAN}" stroke-width="4"/>`;
  for (let gx = 0; gx < 6; gx++) for (let gy = 0; gy < 6; gy++)
    b += `<rect x="${475 + gx * 46}" y="${375 + gy * 46}" width="30" height="30" rx="4" fill="${CYAN}" opacity="${0.25 + r() * 0.6}"/>`;
  [[120, 880], [980, 120], [1060, 860], [150, 130]].forEach(([x, y]) => {
    b += `<rect x="${x}" y="${y}" width="80" height="44" rx="6" fill="#C9D3DA"/><rect x="${x + 22}" y="${y}" width="36" height="44" fill="${STEEL}"/>`;
  });
  return wrap(W, H, b);
}

// 港湾・コンテナ・船
function port() {
  const W = 1600, H = 1000, r = rng(11);
  let b = `<rect width="${W}" height="${H}" fill="url(#sky)"/>`;
  b += `<rect x="0" y="640" width="${W}" height="360" fill="#6FAEBB"/>`;
  for (let i = 0; i < 16; i++) b += `<rect x="${r() * W}" y="${680 + r() * 300}" width="${90 + r() * 150}" height="5" rx="2" fill="#FFFFFF" opacity="0.3"/>`;
  // 岸壁
  b += `<rect x="0" y="560" width="${W}" height="80" fill="${NAVY2}"/>`;
  // クレーン
  [260, 760, 1260].forEach((cx) => {
    b += `<rect x="${cx - 70}" y="250" width="20" height="310" fill="${TEAL}"/><rect x="${cx + 50}" y="250" width="20" height="310" fill="${TEAL}"/>`;
    b += `<rect x="${cx - 160}" y="230" width="380" height="26" fill="${TEAL}"/>`;
    b += `<line x1="${cx - 60}" y1="250" x2="${cx + 60}" y2="400" stroke="${TEAL}" stroke-width="10"/><line x1="${cx + 60}" y1="250" x2="${cx - 60}" y2="400" stroke="${TEAL}" stroke-width="10"/>`;
    b += `<rect x="${cx + 130}" y="256" width="6" height="140" fill="${NAVY}"/><rect x="${cx + 98}" y="396" width="70" height="34" fill="${ORANGE}"/>`;
  });
  // コンテナの山
  const cols = [ORANGE, TEAL, NAVY3, CYAN, "#C9D3DA", NAVY];
  for (let row = 0; row < 3; row++) for (let i = 0; i < 9; i++) {
    if (r() < 0.2 && row > 0) continue;
    const x = 30 + i * 170, y = 520 - row * 44;
    b += `<rect x="${x}" y="${y}" width="160" height="40" fill="${cols[Math.floor(r() * cols.length)]}"/>`;
    for (let k = 1; k < 8; k++) b += `<rect x="${x + k * 20}" y="${y + 4}" width="2" height="32" fill="#000" opacity="0.12"/>`;
  }
  // 船
  b += `<path d="M 220 700 L 1380 700 L 1300 860 L 300 860 Z" fill="${NAVY}"/>`;
  b += `<rect x="220" y="688" width="1160" height="18" fill="#C9D3DA"/>`;
  b += `<rect x="1140" y="560" width="160" height="130" fill="#FFFFFF"/><rect x="1160" y="585" width="120" height="22" fill="${CYAN}"/>`;
  for (let i = 0; i < 6; i++) b += `<rect x="${300 + i * 135}" y="${648}" width="125" height="40" fill="${cols[(i + 2) % cols.length]}"/>`;
  return wrap(W, H, b, skyDefs("sky", "#CFE6EC", "#F4FAFB"));
}

// 鋼板コイル
function coil() {
  const W = 1200, H = 1000;
  const grad = `<linearGradient id="steel" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#E9EEF1"/><stop offset="0.45" stop-color="#AEB9C2"/><stop offset="1" stop-color="#6E7C88"/></linearGradient>
  <radialGradient id="face" cx="0.5" cy="0.5" r="0.5"><stop offset="0" stop-color="#5D6B77"/><stop offset="0.35" stop-color="#DDE4E8"/><stop offset="1" stop-color="#9AA7B1"/></radialGradient>
  ${skyDefs("bg", "#DCEBEF", "#F6FAFB")}`;
  let b = `<rect width="${W}" height="${H}" fill="url(#bg)"/><rect x="0" y="760" width="${W}" height="240" fill="${NAVY2}"/>`;
  b += `<rect x="0" y="760" width="${W}" height="14" fill="${TEAL}"/>`;
  const one = (cx, cy, R) => {
    let s = `<rect x="${cx - R * 1.3}" y="${cy - R}" width="${R * 1.3}" height="${2 * R}" fill="url(#steel)"/>`;
    s += `<ellipse cx="${cx - R * 1.3}" cy="${cy}" rx="${R * 0.42}" ry="${R}" fill="#7E8B96"/>`;
    s += `<ellipse cx="${cx}" cy="${cy}" rx="${R * 0.42}" ry="${R}" fill="url(#face)"/>`;
    for (let k = 1; k < 7; k++) s += `<ellipse cx="${cx}" cy="${cy}" rx="${R * 0.42 * (1 - k * 0.09)}" ry="${R * (1 - k * 0.09)}" fill="none" stroke="#FFFFFF" stroke-opacity="0.35" stroke-width="2"/>`;
    s += `<ellipse cx="${cx}" cy="${cy}" rx="${R * 0.15}" ry="${R * 0.36}" fill="${NAVY}"/>`;
    return s;
  };
  b += one(560, 560, 210);
  b += one(1050, 600, 170);
  b += `<rect x="80" y="700" width="300" height="60" fill="${NAVY3}"/><rect x="100" y="640" width="260" height="60" fill="${STEEL}"/>`;
  return wrap(W, H, b, grad);
}

// 水素還元製鉄のイメージ（再エネ・水素タンク・シャフト炉）
function hydrogen() {
  const W = 1600, H = 1000;
  let b = `<rect width="${W}" height="${H}" fill="url(#sky)"/>`;
  b += `<rect x="0" y="780" width="${W}" height="220" fill="${NAVY}"/>`;
  b += `<rect x="0" y="760" width="${W}" height="24" fill="#5E9EA9"/>`;
  // 風車
  [[150, 360], [330, 300], [1420, 340]].forEach(([x, y]) => {
    b += `<polygon points="${x - 8},${y} ${x + 8},${y} ${x + 14},760 ${x - 14},760" fill="#FFFFFF"/>`;
    [0, 120, 240].forEach((a) => {
      b += `<rect x="${x - 6}" y="${y - 170}" width="12" height="170" rx="6" fill="#FFFFFF" transform="rotate(${a + 20} ${x} ${y})"/>`;
    });
    b += `<circle cx="${x}" cy="${y}" r="14" fill="${CYAN}"/>`;
  });
  // 太陽光パネル
  for (let i = 0; i < 5; i++) b += `<polygon points="${470 + i * 70},700 ${530 + i * 70},700 ${510 + i * 70},650 ${450 + i * 70},650" fill="${NAVY3}" stroke="${CYAN}" stroke-width="3"/>`;
  // 水素タンク
  [[860, 600], [990, 600]].forEach(([x, y]) => {
    b += `<rect x="${x - 50}" y="${y - 160}" width="100" height="320" rx="50" fill="#FFFFFF"/>`;
    b += `<rect x="${x - 50}" y="${y - 40}" width="100" height="26" fill="${CYAN}"/>`;
  });
  // シャフト炉
  b += `<rect x="1120" y="250" width="150" height="530" fill="${TEAL}"/>`;
  b += `<rect x="1100" y="230" width="190" height="40" fill="${NAVY2}"/>`;
  b += `<rect x="1150" y="140" width="40" height="90" fill="${NAVY2}"/>`;
  for (let k = 0; k < 5; k++) b += `<rect x="1120" y="${320 + k * 90}" width="150" height="10" fill="${NAVY2}" opacity="0.5"/>`;
  b += `<rect x="1270" y="560" width="120" height="220" fill="${NAVY2}"/>`;
  // 配管
  b += `<path d="M 1040 520 L 1120 520" stroke="${CYAN}" stroke-width="18"/><path d="M 990 440 L 990 400 L 1120 400" stroke="${CYAN}" stroke-width="14" fill="none"/>`;
  b += `<path d="M 640 650 L 640 560 L 810 560" stroke="${ORANGE}" stroke-width="10" fill="none"/>`;
  // 蒸気（水）
  b += `<ellipse cx="1180" cy="100" rx="80" ry="36" fill="#FFFFFF" opacity="0.8"/><ellipse cx="1260" cy="60" rx="70" ry="30" fill="#FFFFFF" opacity="0.55"/>`;
  return wrap(W, H, b, skyDefs("sky", "#CBE7EC", "#F3FAFB"));
}

// 生活で目にする商品（食品・化粧品・アプリ・ゲーム）
function consumer() {
  const W = 1600, H = 1000;
  let b = `<rect width="${W}" height="${H}" fill="url(#bg)"/>`;
  b += `<rect x="0" y="760" width="${W}" height="240" fill="#F2D6B0"/><rect x="0" y="760" width="${W}" height="16" fill="${ORANGE}"/>`;
  // 飲料ボトル
  b += `<rect x="150" y="380" width="150" height="380" rx="40" fill="${CYAN}"/><rect x="190" y="320" width="70" height="70" rx="10" fill="${CYAN}"/><rect x="185" y="300" width="80" height="30" rx="8" fill="${NAVY2}"/>`;
  b += `<rect x="150" y="500" width="150" height="110" fill="#FFFFFF" opacity="0.85"/>`;
  // 菓子の箱
  b += `<rect x="340" y="520" width="210" height="240" rx="12" fill="${ORANGE}"/><circle cx="445" cy="630" r="60" fill="#FFFFFF" opacity="0.85"/>`;
  // 化粧品
  b += `<rect x="600" y="560" width="110" height="200" rx="18" fill="#F7B9C4"/><rect x="620" y="500" width="70" height="70" rx="8" fill="#FFFFFF"/>`;
  b += `<ellipse cx="800" cy="740" rx="70" ry="22" fill="#E58FA0"/><rect x="730" y="680" width="140" height="60" fill="#F7B9C4"/><ellipse cx="800" cy="680" rx="70" ry="22" fill="#FBD5DC"/>`;
  // スマホとアプリ
  b += `<rect x="950" y="260" width="280" height="500" rx="40" fill="${NAVY}"/><rect x="972" y="300" width="236" height="420" rx="18" fill="#FFFFFF"/>`;
  const appCols = [ORANGE, CYAN, TEAL, "#F7B9C4", NAVY3, "#A7D7A0"];
  for (let i = 0; i < 12; i++) {
    const cx = 1000 + (i % 3) * 70, cy = 330 + Math.floor(i / 3) * 80;
    b += `<rect x="${cx}" y="${cy}" width="50" height="50" rx="12" fill="${appCols[i % appCols.length]}"/>`;
  }
  // ゲームコントローラー
  b += `<path d="M 1290 640 Q 1300 560 1380 560 L 1500 560 Q 1580 560 1590 640 L 1600 720 Q 1600 770 1560 760 L 1500 720 L 1380 720 L 1320 760 Q 1280 770 1280 720 Z" fill="${NAVY3}"/>`;
  b += `<rect x="1340" y="620" width="50" height="16" fill="#FFFFFF"/><rect x="1357" y="603" width="16" height="50" fill="#FFFFFF"/>`;
  b += `<circle cx="1510" cy="615" r="12" fill="${ORANGE}"/><circle cx="1540" cy="645" r="12" fill="${CYAN}"/>`;
  return wrap(W, H, b, skyDefs("bg", "#FFF3E2", "#FFFFFF"));
}

// 高炉の模式図（透明背景）
function blastFurnace() {
  const W = 600, H = 800;
  let b = `<path d="M 170 760 L 190 300 L 240 140 L 360 140 L 410 300 L 430 760 Z" fill="${NAVY}"/>`;
  b += `<rect x="225" y="60" width="150" height="80" fill="${NAVY3}"/>`;
  b += `<path d="M 205 300 L 395 300 L 410 560 L 190 560 Z" fill="${NAVY3}"/>`;
  for (let i = 0; i < 6; i++) b += `<circle cx="${230 + (i % 3) * 70}" cy="${340 + Math.floor(i / 3) * 70}" r="22" fill="#B5643C" opacity="0.9"/>`;
  b += `<rect x="195" y="600" width="210" height="140" fill="${ORANGE}"/>`;
  b += `<ellipse cx="300" cy="600" rx="105" ry="20" fill="#F7C987"/>`;
  b += `<path d="M 430 700 L 560 700" stroke="${ORANGE}" stroke-width="22"/>`;
  b += `<path d="M 140 640 L 195 640 M 140 680 L 195 680" stroke="${CYAN}" stroke-width="14"/>`;
  return wrap(W, H, b);
}

// 電炉の模式図（透明背景）
function electricFurnace() {
  const W = 600, H = 800;
  let b = "";
  [230, 300, 370].forEach((x) => {
    b += `<rect x="${x - 16}" y="90" width="32" height="360" fill="#5D6B77"/>`;
  });
  b += `<path d="M 330 40 L 280 150 L 320 150 L 270 260 L 380 130 L 335 130 L 380 40 Z" fill="${ORANGE}"/>`;
  b += `<path d="M 110 400 L 490 400 L 470 470 L 130 470 Z" fill="${NAVY3}"/>`;
  b += `<path d="M 120 470 L 480 470 L 450 720 Q 300 780 150 720 Z" fill="${NAVY}"/>`;
  b += `<path d="M 160 520 L 440 520 L 425 680 Q 300 720 175 680 Z" fill="${ORANGE}"/>`;
  b += `<ellipse cx="300" cy="520" rx="140" ry="18" fill="#F7C987"/>`;
  for (let i = 0; i < 5; i++) b += `<rect x="${190 + i * 48}" y="${560 + (i % 2) * 40}" width="40" height="22" rx="4" fill="#9AA7B1" transform="rotate(${i * 25 - 40} ${210 + i * 48} ${570})"/>`;
  return wrap(W, H, b);
}

// 金属の質感を表す球
function metalSphere(base, light, dark) {
  return wrap(400, 400,
    `<circle cx="200" cy="200" r="180" fill="url(#g)"/><ellipse cx="140" cy="120" rx="60" ry="34" fill="#FFFFFF" opacity="0.55" transform="rotate(-30 140 120)"/>`,
    `<radialGradient id="g" cx="0.35" cy="0.3" r="0.8"><stop offset="0" stop-color="${light}"/><stop offset="0.5" stop-color="${base}"/><stop offset="1" stop-color="${dark}"/></radialGradient>`);
}

// グループディスカッションの様子（人物は抽象的なシルエット）
function discussion() {
  const W = 1200, H = 1000;
  let b = `<rect width="${W}" height="${H}" fill="url(#bg)"/>`;
  // 壁のボードと付箋
  b += `<rect x="160" y="110" width="880" height="380" rx="20" fill="#FFFFFF"/>`;
  const notes = [ORANGE, CYAN, "#FBE29A", "#A7D7A0", ORANGE, "#F7B9C4", CYAN, "#FBE29A", "#A7D7A0"];
  notes.forEach((c, i) => {
    const x = 220 + (i % 5) * 160, y = 160 + Math.floor(i / 5) * 160;
    b += `<rect x="${x}" y="${y}" width="120" height="110" fill="${c}" transform="rotate(${(i % 3) - 1} ${x + 60} ${y + 55})"/>`;
    b += `<rect x="${x + 18}" y="${y + 30}" width="80" height="8" fill="#000" opacity="0.15"/><rect x="${x + 18}" y="${y + 55}" width="60" height="8" fill="#000" opacity="0.15"/>`;
  });
  b += `<line x1="380" y1="215" x2="520" y2="215" stroke="${NAVY3}" stroke-width="4" stroke-dasharray="10 8"/>`;
  // 人物
  const person = (x, y, col, s = 1) =>
    `<circle cx="${x}" cy="${y}" r="${48 * s}" fill="#F2C9A5"/><path d="M ${x - 48 * s} ${y - 8 * s} Q ${x} ${y - 80 * s} ${x + 48 * s} ${y - 8 * s} Q ${x + 30 * s} ${y - 40 * s} ${x} ${y - 38 * s} Q ${x - 30 * s} ${y - 40 * s} ${x - 48 * s} ${y - 8 * s} Z" fill="${NAVY}"/>` +
    `<rect x="${x - 85 * s}" y="${y + 58 * s}" width="${170 * s}" height="${200 * s}" rx="${70 * s}" fill="${col}"/>`;
  b += person(250, 560, TEAL);
  b += person(480, 540, ORANGE);
  b += person(720, 540, NAVY3);
  b += person(950, 560, CYAN);
  // 机
  b += `<rect x="80" y="760" width="1040" height="60" rx="16" fill="${NAVY2}"/><rect x="140" y="820" width="40" height="180" fill="${NAVY2}"/><rect x="1020" y="820" width="40" height="180" fill="${NAVY2}"/>`;
  b += `<rect x="420" y="700" width="220" height="64" rx="8" fill="#C9D3DA"/><rect x="760" y="730" width="120" height="30" fill="#FFFFFF"/>`;
  return wrap(W, H, b, skyDefs("bg", "#DCEFF3", "#F6FBFC"));
}

module.exports = {
  cityBridge, steelMill, chip, port, coil, hydrogen, consumer,
  blastFurnace, electricFurnace, metalSphere, discussion,
};
