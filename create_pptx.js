const PptxGenJS = require("./node_modules/pptxgenjs");

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE"; // 16:9 13.33" x 7.5"

const W = 13.33;
const H = 7.5;

// Color palette
const C = {
  teal: "028090",
  tealLight: "E0F4F7",
  tealMid: "A8D8DF",
  white: "FFFFFF",
  textDark: "1A1A2E",
  textMuted: "6B7280",
  orange: "F97316",
  codeBg: "F3F4F6",
  codeText: "1F2937",
  lightBg: "F8FAFB",
};

// ── Helper: header bar ────────────────────────────────────────────────────────
function addHeader(slide, title) {
  // Teal header bar
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: W, h: 1.05,
    fill: { color: C.teal },
    line: { color: C.teal },
  });
  slide.addText(title, {
    x: 0.45, y: 0.12, w: W - 0.9, h: 0.82,
    fontSize: 28,
    bold: true,
    color: C.white,
    fontFace: "Calibri",
    valign: "middle",
  });
  // Decorative accent dot (orange)
  slide.addShape(pptx.ShapeType.ellipse, {
    x: W - 0.55, y: 0.28, w: 0.22, h: 0.22,
    fill: { color: C.orange },
    line: { color: C.orange },
  });
}

// ── Helper: footer ────────────────────────────────────────────────────────────
function addFooter(slide, pageNum) {
  slide.addText(`Usutakuの講義  |  Claude Code 入門`, {
    x: 0.45, y: H - 0.38, w: W - 2, h: 0.28,
    fontSize: 9,
    color: C.textMuted,
    fontFace: "Calibri",
  });
  slide.addText(`${pageNum} / 8`, {
    x: W - 1.1, y: H - 0.38, w: 0.8, h: 0.28,
    fontSize: 9,
    color: C.textMuted,
    fontFace: "Calibri",
    align: "right",
  });
  // Thin bottom line
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: H - 0.45, w: W, h: 0.02,
    fill: { color: C.tealMid },
    line: { color: C.tealMid },
  });
}

// ── Helper: bullet row with colored icon circle ───────────────────────────────
function addBulletRow(slide, iconChar, label, desc, x, y, iconColor) {
  const ic = iconColor || C.teal;
  // Icon circle
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x, y: y + 0.02, w: 0.38, h: 0.38,
    fill: { color: ic },
    line: { color: ic },
  });
  slide.addText(iconChar, {
    x: x, y: y + 0.02, w: 0.38, h: 0.38,
    fontSize: 14,
    bold: true,
    color: C.white,
    fontFace: "Calibri",
    align: "center",
    valign: "middle",
  });
  // Label
  slide.addText(label, {
    x: x + 0.5, y: y, w: 4.5, h: 0.22,
    fontSize: 13,
    bold: true,
    color: C.textDark,
    fontFace: "Calibri",
  });
  // Description
  slide.addText(desc, {
    x: x + 0.5, y: y + 0.22, w: 4.5, h: 0.2,
    fontSize: 10.5,
    color: C.textMuted,
    fontFace: "Calibri",
  });
}

// ── Helper: code block ────────────────────────────────────────────────────────
function addCodeBlock(slide, code, x, y, w, h) {
  slide.addShape(pptx.ShapeType.rect, {
    x: x, y: y, w: w, h: h,
    fill: { color: C.codeBg },
    line: { color: "D1D5DB", width: 1 },
  });
  slide.addText(code, {
    x: x + 0.18, y: y + 0.08, w: w - 0.36, h: h - 0.16,
    fontSize: 12.5,
    color: C.codeText,
    fontFace: "Courier New",
    valign: "top",
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 1: Cover
// ════════════════════════════════════════════════════════════════════════════
const s1 = pptx.addSlide();
s1.background = { color: C.white };

// Left panel background
s1.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: W * 0.56, h: H,
  fill: { color: C.white },
  line: { color: C.white },
});

// Right panel — teal
s1.addShape(pptx.ShapeType.rect, {
  x: W * 0.56, y: 0, w: W * 0.44, h: H,
  fill: { color: C.teal },
  line: { color: C.teal },
});

// Decorative shapes on right panel
s1.addShape(pptx.ShapeType.rect, {
  x: W * 0.65, y: 0.6, w: 2.4, h: 2.4,
  fill: { color: "026E7C" },
  line: { color: "026E7C" },
});
s1.addShape(pptx.ShapeType.rect, {
  x: W * 0.72, y: 1.4, w: 2.1, h: 2.1,
  fill: { color: "03A4B8" },
  line: { color: "03A4B8" },
});
s1.addShape(pptx.ShapeType.ellipse, {
  x: W * 0.74, y: 3.8, w: 1.5, h: 1.5,
  fill: { color: C.orange },
  line: { color: C.orange },
});
s1.addShape(pptx.ShapeType.rect, {
  x: W * 0.58, y: 5.5, w: 3.5, h: 0.06,
  fill: { color: "02B4CC" },
  line: { color: "02B4CC" },
});

// Teal top-left accent bar
s1.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: 0.18, h: H,
  fill: { color: C.teal },
  line: { color: C.teal },
});

// Badge: "入門"
s1.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.38,
  fill: { color: C.tealLight },
  line: { color: C.teal, width: 1.2 },
});
s1.addText("入門講義", {
  x: 0.5, y: 1.0, w: 1.2, h: 0.38,
  fontSize: 12,
  bold: true,
  color: C.teal,
  fontFace: "Calibri",
  align: "center",
  valign: "middle",
});

// Main title
s1.addText("Claude Code", {
  x: 0.48, y: 1.55, w: 6.5, h: 1.1,
  fontSize: 54,
  bold: true,
  color: C.textDark,
  fontFace: "Calibri",
});
s1.addText("はじめての一歩", {
  x: 0.48, y: 2.6, w: 6.5, h: 0.7,
  fontSize: 28,
  color: C.teal,
  fontFace: "Calibri",
});

// Divider
s1.addShape(pptx.ShapeType.rect, {
  x: 0.48, y: 3.42, w: 3.2, h: 0.05,
  fill: { color: C.tealMid },
  line: { color: C.tealMid },
});

// Subtitle info
s1.addText([
  { text: "Usutaku", options: { bold: true, color: C.textDark, fontSize: 14 } },
  { text: "  |  2026年4月", options: { color: C.textMuted, fontSize: 13 } },
], {
  x: 0.48, y: 3.6, w: 6.5, h: 0.38,
  fontFace: "Calibri",
});

s1.addText("AIを使いこなして、開発をもっとスマートに", {
  x: 0.48, y: 4.1, w: 6.5, h: 0.35,
  fontSize: 12,
  color: C.textMuted,
  fontFace: "Calibri",
  italics: true,
});

// Bottom brand bar
s1.addShape(pptx.ShapeType.rect, {
  x: 0, y: H - 0.42, w: W * 0.56, h: 0.42,
  fill: { color: C.tealLight },
  line: { color: C.tealLight },
});
s1.addText("Anthropic Claude Code  ·  CLI Tool", {
  x: 0.48, y: H - 0.42, w: W * 0.54, h: 0.42,
  fontSize: 9.5,
  color: C.teal,
  fontFace: "Calibri",
  valign: "middle",
});

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 2: Claude Codeとは？
// ════════════════════════════════════════════════════════════════════════════
const s2 = pptx.addSlide();
s2.background = { color: C.lightBg };
addHeader(s2, "Claude Code とは？");
addFooter(s2, 2);

// Main description box
s2.addShape(pptx.ShapeType.rect, {
  x: 0.45, y: 1.25, w: W - 0.9, h: 0.85,
  fill: { color: C.tealLight },
  line: { color: C.tealMid, width: 1 },
});
s2.addText([
  { text: "Claude Code", options: { bold: true, color: C.teal } },
  { text: " は Anthropic が開発した ", options: { color: C.textDark } },
  { text: "AI 搭載のコマンドラインツール (CLI)", options: { bold: true, color: C.textDark } },
  { text: " です。\n自然言語で指示するだけで、コードの生成・編集・説明・デバッグを自動で行います。", options: { color: C.textDark } },
], {
  x: 0.65, y: 1.28, w: W - 1.3, h: 0.78,
  fontSize: 13.5,
  fontFace: "Calibri",
  valign: "middle",
});

// Feature cards: 2 rows × 2 cols
const features = [
  { icon: "✏", label: "コード生成", desc: "関数・クラス・テストを自動生成", color: C.teal },
  { icon: "🔍", label: "コード説明", desc: "難しいコードをわかりやすく解説", color: "6366F1" },
  { icon: "🐛", label: "バグ修正", desc: "エラーの原因を特定して自動修正", color: C.orange },
  { icon: "📁", label: "ファイル操作", desc: "ファイルの読み書き・リファクタ", color: "10B981" },
];

const cardW = 5.8;
const cardH = 1.35;
const positions = [
  { x: 0.45, y: 2.3 }, { x: 6.9, y: 2.3 },
  { x: 0.45, y: 3.82 }, { x: 6.9, y: 3.82 },
];
features.forEach((f, i) => {
  const p = positions[i];
  s2.addShape(pptx.ShapeType.rect, {
    x: p.x, y: p.y, w: cardW, h: cardH,
    fill: { color: C.white },
    line: { color: "E5E7EB", width: 1 },
  });
  // Color stripe
  s2.addShape(pptx.ShapeType.rect, {
    x: p.x, y: p.y, w: 0.12, h: cardH,
    fill: { color: f.color },
    line: { color: f.color },
  });
  s2.addText(f.icon, {
    x: p.x + 0.22, y: p.y + 0.22, w: 0.7, h: 0.7,
    fontSize: 26,
    align: "center",
  });
  s2.addText(f.label, {
    x: p.x + 1.05, y: p.y + 0.2, w: cardW - 1.2, h: 0.38,
    fontSize: 15,
    bold: true,
    color: C.textDark,
    fontFace: "Calibri",
  });
  s2.addText(f.desc, {
    x: p.x + 1.05, y: p.y + 0.62, w: cardW - 1.2, h: 0.38,
    fontSize: 11.5,
    color: C.textMuted,
    fontFace: "Calibri",
  });
});

// Bottom note
s2.addText("🤖  ターミナル上で動作するため、IDE不要。既存プロジェクトにすぐ組み込めます。", {
  x: 0.45, y: 5.36, w: W - 0.9, h: 0.36,
  fontSize: 11,
  color: C.teal,
  fontFace: "Calibri",
  bold: true,
});

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 3: インストール
// ════════════════════════════════════════════════════════════════════════════
const s3 = pptx.addSlide();
s3.background = { color: C.white };
addHeader(s3, "インストール");
addFooter(s3, 3);

// Steps
const steps = [
  { num: "1", title: "Node.js をインストール", code: "https://nodejs.org  （v18 以上推奨）", note: "まだインストールしていない場合のみ" },
  { num: "2", title: "Claude Code をインストール", code: "npm install -g @anthropic-ai/claude-code", note: "ターミナルで実行（macOS / Linux / Windows 対応）" },
  { num: "3", title: "APIキーを設定して起動", code: "claude", note: "初回起動時に Anthropic のアカウント認証を行います" },
];

steps.forEach((step, i) => {
  const y = 1.25 + i * 1.73;
  // Step number badge
  s3.addShape(pptx.ShapeType.rect, {
    x: 0.45, y: y, w: 0.55, h: 0.55,
    fill: { color: C.teal },
    line: { color: C.teal },
  });
  s3.addText(step.num, {
    x: 0.45, y: y, w: 0.55, h: 0.55,
    fontSize: 20,
    bold: true,
    color: C.white,
    fontFace: "Calibri",
    align: "center",
    valign: "middle",
  });
  // Title
  s3.addText(step.title, {
    x: 1.15, y: y, w: W - 1.6, h: 0.45,
    fontSize: 15,
    bold: true,
    color: C.textDark,
    fontFace: "Calibri",
    valign: "middle",
  });
  // Code block
  addCodeBlock(s3, step.code, 1.15, y + 0.5, W - 1.6, 0.52);
  // Note
  s3.addText("💡  " + step.note, {
    x: 1.15, y: y + 1.09, w: W - 1.6, h: 0.24,
    fontSize: 10,
    color: C.textMuted,
    fontFace: "Calibri",
  });
  // Connector arrow (between steps)
  if (i < 2) {
    s3.addShape(pptx.ShapeType.rect, {
      x: 0.63, y: y + 1.5, w: 0.02, h: 0.14,
      fill: { color: C.tealMid },
      line: { color: C.tealMid },
    });
  }
});

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 4: 起動と基本UI
// ════════════════════════════════════════════════════════════════════════════
const s4 = pptx.addSlide();
s4.background = { color: C.lightBg };
addHeader(s4, "起動と基本 UI");
addFooter(s4, 4);

// Left: startup explanation
s4.addText("起動コマンド", {
  x: 0.45, y: 1.28, w: 5.8, h: 0.38,
  fontSize: 15,
  bold: true,
  color: C.textDark,
  fontFace: "Calibri",
});
addCodeBlock(s4, "$ claude", 0.45, 1.72, 5.8, 0.52);

s4.addText("プロジェクトを指定して起動", {
  x: 0.45, y: 2.38, w: 5.8, h: 0.38,
  fontSize: 15,
  bold: true,
  color: C.textDark,
  fontFace: "Calibri",
});
addCodeBlock(s4, "$ claude --project /path/to/project", 0.45, 2.82, 5.8, 0.52);

// Description
s4.addText("起動すると対話型の画面が表示されます。\n日本語でそのまま質問・指示を入力できます。", {
  x: 0.45, y: 3.52, w: 5.8, h: 0.62,
  fontSize: 12.5,
  color: C.textMuted,
  fontFace: "Calibri",
});

// Right: UI diagram
s4.addShape(pptx.ShapeType.rect, {
  x: 6.95, y: 1.25, w: 5.88, h: 4.7,
  fill: { color: "1E1E2E" },
  line: { color: "3A3A4A", width: 1.5 },
});
// Terminal header bar
s4.addShape(pptx.ShapeType.rect, {
  x: 6.95, y: 1.25, w: 5.88, h: 0.42,
  fill: { color: "2A2A3E" },
  line: { color: "2A2A3E" },
});
s4.addText("● ● ●", {
  x: 7.1, y: 1.29, w: 1.2, h: 0.3,
  fontSize: 11,
  color: "6B7280",
  fontFace: "Courier New",
});
s4.addText("Terminal", {
  x: 9.6, y: 1.29, w: 3.0, h: 0.3,
  fontSize: 10,
  color: "9CA3AF",
  fontFace: "Calibri",
  align: "center",
});
// Terminal content
const termLines = [
  { text: "$ claude", color: "A8D8DF" },
  { text: "", color: "FFFFFF" },
  { text: "  ╔══════════════════╗", color: "028090" },
  { text: "  ║   Claude Code    ║", color: "028090" },
  { text: "  ╚══════════════════╝", color: "028090" },
  { text: "", color: "FFFFFF" },
  { text: "  How can I help?", color: "E2E8F0" },
  { text: "", color: "FFFFFF" },
  { text: "  > |", color: "F97316" },
];
termLines.forEach((line, idx) => {
  s4.addText(line.text, {
    x: 7.1, y: 1.75 + idx * 0.36, w: 5.5, h: 0.34,
    fontSize: 11.5,
    color: line.color,
    fontFace: "Courier New",
  });
});

// Tips
s4.addText("💡  ヒント", {
  x: 0.45, y: 4.28, w: 5.8, h: 0.3,
  fontSize: 12,
  bold: true,
  color: C.orange,
  fontFace: "Calibri",
});
s4.addText("終了するには Ctrl + C または /exit と入力してください。", {
  x: 0.45, y: 4.58, w: 5.8, h: 0.28,
  fontSize: 11.5,
  color: C.textMuted,
  fontFace: "Calibri",
});
s4.addText("会話は自動的に保存されるため、再起動後も続きから使えます。", {
  x: 0.45, y: 4.86, w: 5.8, h: 0.28,
  fontSize: 11.5,
  color: C.textMuted,
  fontFace: "Calibri",
});

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 5: 基本コマンド（スラッシュコマンド）
// ════════════════════════════════════════════════════════════════════════════
const s5 = pptx.addSlide();
s5.background = { color: C.white };
addHeader(s5, "基本コマンド（スラッシュコマンド）");
addFooter(s5, 5);

s5.addText("Claude Code では / で始まるコマンドを使って動作をコントロールできます。", {
  x: 0.45, y: 1.18, w: W - 0.9, h: 0.35,
  fontSize: 12.5,
  color: C.textMuted,
  fontFace: "Calibri",
});

const cmds = [
  { cmd: "/help", desc: "使えるコマンドの一覧を表示", color: C.teal },
  { cmd: "/clear", desc: "会話履歴をリセット（新しいタスク開始時に便利）", color: "6366F1" },
  { cmd: "/compact", desc: "長い会話を要約して文脈を圧縮する", color: "10B981" },
  { cmd: "/exit", desc: "Claude Code を終了する", color: C.orange },
  { cmd: "/status", desc: "現在の設定・モデル・使用状況を確認", color: "8B5CF6" },
  { cmd: "/config", desc: "設定の確認・変更（モデル切り替えなど）", color: "EC4899" },
];

const colW = 5.8;
const rowH = 0.78;
cmds.forEach((c, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.45 + col * 6.44;
  const y = 1.65 + row * (rowH + 0.14);

  s5.addShape(pptx.ShapeType.rect, {
    x: x, y: y, w: colW, h: rowH,
    fill: { color: C.white },
    line: { color: "E5E7EB", width: 1 },
  });
  s5.addShape(pptx.ShapeType.rect, {
    x: x, y: y, w: 0.1, h: rowH,
    fill: { color: c.color },
    line: { color: c.color },
  });
  s5.addText(c.cmd, {
    x: x + 0.22, y: y + 0.1, w: colW - 0.32, h: 0.3,
    fontSize: 14.5,
    bold: true,
    color: c.color,
    fontFace: "Courier New",
  });
  s5.addText(c.desc, {
    x: x + 0.22, y: y + 0.42, w: colW - 0.32, h: 0.26,
    fontSize: 10.5,
    color: C.textMuted,
    fontFace: "Calibri",
  });
});

// Tip
s5.addShape(pptx.ShapeType.rect, {
  x: 0.45, y: 5.4, w: W - 0.9, h: 0.5,
  fill: { color: C.tealLight },
  line: { color: C.tealMid, width: 1 },
});
s5.addText("💡  /help と入力すると、インストール済みスキルや全コマンドを一覧表示できます。まず /help から試してみましょう！", {
  x: 0.65, y: 5.41, w: W - 1.3, h: 0.48,
  fontSize: 11.5,
  color: C.teal,
  fontFace: "Calibri",
  bold: true,
  valign: "middle",
});

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 6: 効果的な使い方
// ════════════════════════════════════════════════════════════════════════════
const s6 = pptx.addSlide();
s6.background = { color: C.lightBg };
addHeader(s6, "効果的な使い方");
addFooter(s6, 6);

// Left column: bad vs good
s6.addText("指示の出し方が大切です", {
  x: 0.45, y: 1.2, w: 6.0, h: 0.38,
  fontSize: 16,
  bold: true,
  color: C.textDark,
  fontFace: "Calibri",
});

// Bad example
s6.addShape(pptx.ShapeType.rect, {
  x: 0.45, y: 1.68, w: 0.7, h: 0.38,
  fill: { color: "FEE2E2" },
  line: { color: "FCA5A5", width: 1 },
});
s6.addText("NG", {
  x: 0.45, y: 1.68, w: 0.7, h: 0.38,
  fontSize: 13,
  bold: true,
  color: "DC2626",
  fontFace: "Calibri",
  align: "center",
  valign: "middle",
});
addCodeBlock(s6, "バグを直して", 1.25, 1.68, 5.2, 0.44);

// Good example
s6.addShape(pptx.ShapeType.rect, {
  x: 0.45, y: 2.25, w: 0.7, h: 0.38,
  fill: { color: "D1FAE5" },
  line: { color: "6EE7B7", width: 1 },
});
s6.addText("OK", {
  x: 0.45, y: 2.25, w: 0.7, h: 0.38,
  fontSize: 13,
  bold: true,
  color: "059669",
  fontFace: "Calibri",
  align: "center",
  valign: "middle",
});
addCodeBlock(s6,
  "src/auth.py の login() 関数で\nTypeError が発生しています。\n原因を調べて修正してください。",
  1.25, 2.25, 5.2, 0.88);

// Tips list
const tips = [
  { num: "1", text: "ファイル名・関数名を具体的に伝える" },
  { num: "2", text: "エラーメッセージをそのまま貼り付ける" },
  { num: "3", text: "「何をしたいか」だけでなく「なぜか」も伝える" },
  { num: "4", text: "大きなタスクは小さく分割して依頼する" },
];
s6.addText("コツ", {
  x: 0.45, y: 3.32, w: 5.8, h: 0.35,
  fontSize: 14,
  bold: true,
  color: C.teal,
  fontFace: "Calibri",
});
tips.forEach((t, i) => {
  s6.addShape(pptx.ShapeType.ellipse, {
    x: 0.45, y: 3.74 + i * 0.48, w: 0.28, h: 0.28,
    fill: { color: C.teal },
    line: { color: C.teal },
  });
  s6.addText(t.num, {
    x: 0.45, y: 3.74 + i * 0.48, w: 0.28, h: 0.28,
    fontSize: 11,
    bold: true,
    color: C.white,
    fontFace: "Calibri",
    align: "center",
    valign: "middle",
  });
  s6.addText(t.text, {
    x: 0.88, y: 3.72 + i * 0.48, w: 5.2, h: 0.32,
    fontSize: 12,
    color: C.textDark,
    fontFace: "Calibri",
    valign: "middle",
  });
});

// Right column: context panel
s6.addShape(pptx.ShapeType.rect, {
  x: 7.0, y: 1.2, w: 5.83, h: 4.65,
  fill: { color: C.white },
  line: { color: "E5E7EB", width: 1 },
});
s6.addShape(pptx.ShapeType.rect, {
  x: 7.0, y: 1.2, w: 5.83, h: 0.48,
  fill: { color: C.tealLight },
  line: { color: C.tealLight },
});
s6.addText("コンテキスト（文脈）の共有", {
  x: 7.15, y: 1.23, w: 5.5, h: 0.42,
  fontSize: 14,
  bold: true,
  color: C.teal,
  fontFace: "Calibri",
  valign: "middle",
});

const contextItems = [
  { icon: "📋", title: "CLAUDE.md を活用", body: "プロジェクトのルールや仕様を\nCLAUDE.md に書いておくと\n自動的に読み込まれます" },
  { icon: "📂", title: "ファイルを参照", body: "@ファイル名 で特定のファイルを\n会話の文脈に含められます" },
  { icon: "🔄", title: "/compact を使う", body: "長い会話が続いたら /compact で\n要約して文脈をリフレッシュ" },
];

contextItems.forEach((item, i) => {
  const cy = 1.82 + i * 1.3;
  s6.addText(item.icon, {
    x: 7.15, y: cy, w: 0.55, h: 0.55,
    fontSize: 22,
    align: "center",
  });
  s6.addText(item.title, {
    x: 7.78, y: cy, w: 4.8, h: 0.32,
    fontSize: 13,
    bold: true,
    color: C.textDark,
    fontFace: "Calibri",
  });
  s6.addText(item.body, {
    x: 7.78, y: cy + 0.33, w: 4.8, h: 0.72,
    fontSize: 10.5,
    color: C.textMuted,
    fontFace: "Calibri",
  });
  if (i < 2) {
    s6.addShape(pptx.ShapeType.rect, {
      x: 7.15, y: cy + 1.2, w: 5.5, h: 0.01,
      fill: { color: "E5E7EB" },
      line: { color: "E5E7EB" },
    });
  }
});

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 7: よくある使用シーン
// ════════════════════════════════════════════════════════════════════════════
const s7 = pptx.addSlide();
s7.background = { color: C.white };
addHeader(s7, "よくある使用シーン");
addFooter(s7, 7);

const scenes = [
  {
    icon: "🐛",
    title: "バグ修正",
    color: C.orange,
    example: "> src/utils.py の calculate() で\n  ZeroDivisionError が出ます。\n  原因を見つけて修正してください。",
  },
  {
    icon: "⚡",
    title: "コード生成",
    color: C.teal,
    example: "> ユーザーのメールアドレスを\n  検証する Python 関数を書いて。\n  テストコードも含めて。",
  },
  {
    icon: "📖",
    title: "コード説明",
    color: "6366F1",
    example: "> この関数が何をしているか\n  初心者にもわかるように\n  日本語で説明してください。",
  },
  {
    icon: "🔄",
    title: "リファクタリング",
    color: "10B981",
    example: "> auth.js を読みやすくリファクタ\n  してください。変数名も\n  わかりやすくしてほしい。",
  },
];

const sceneW = 5.8;
const sceneH = 2.72;
const scenePos = [
  { x: 0.45, y: 1.25 }, { x: 6.9, y: 1.25 },
  { x: 0.45, y: 4.1 }, { x: 6.9, y: 4.1 },
];

scenes.forEach((sc, i) => {
  const p = scenePos[i];
  s7.addShape(pptx.ShapeType.rect, {
    x: p.x, y: p.y, w: sceneW, h: sceneH,
    fill: { color: C.lightBg },
    line: { color: "E5E7EB", width: 1 },
  });
  // Top color band
  s7.addShape(pptx.ShapeType.rect, {
    x: p.x, y: p.y, w: sceneW, h: 0.52,
    fill: { color: sc.color },
    line: { color: sc.color },
  });
  s7.addText(sc.icon + "  " + sc.title, {
    x: p.x + 0.18, y: p.y, w: sceneW - 0.36, h: 0.52,
    fontSize: 15,
    bold: true,
    color: C.white,
    fontFace: "Calibri",
    valign: "middle",
  });
  // Example code
  s7.addShape(pptx.ShapeType.rect, {
    x: p.x + 0.18, y: p.y + 0.62, w: sceneW - 0.36, h: 1.8,
    fill: { color: C.codeBg },
    line: { color: "D1D5DB", width: 1 },
  });
  s7.addText(sc.example, {
    x: p.x + 0.32, y: p.y + 0.7, w: sceneW - 0.64, h: 1.64,
    fontSize: 10.5,
    color: C.codeText,
    fontFace: "Courier New",
    valign: "top",
  });
});

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8: まとめ
// ════════════════════════════════════════════════════════════════════════════
const s8 = pptx.addSlide();
s8.background = { color: C.teal };

// Decorative shapes
s8.addShape(pptx.ShapeType.rect, {
  x: W - 3.5, y: 0, w: 3.5, h: H,
  fill: { color: "026E7C" },
  line: { color: "026E7C" },
});
s8.addShape(pptx.ShapeType.ellipse, {
  x: W - 2.8, y: 1.2, w: 2.2, h: 2.2,
  fill: { color: "03A4B8" },
  line: { color: "03A4B8" },
});
s8.addShape(pptx.ShapeType.ellipse, {
  x: W - 1.8, y: 4.5, w: 1.4, h: 1.4,
  fill: { color: C.orange },
  line: { color: C.orange },
});

// Title area
s8.addText("まとめ", {
  x: 0.55, y: 0.6, w: 8.5, h: 0.62,
  fontSize: 14,
  color: "A8D8DF",
  fontFace: "Calibri",
  bold: true,
});
s8.addText("Claude Code を使いこなそう", {
  x: 0.55, y: 1.15, w: 8.5, h: 0.95,
  fontSize: 36,
  bold: true,
  color: C.white,
  fontFace: "Calibri",
});

// Divider
s8.addShape(pptx.ShapeType.rect, {
  x: 0.55, y: 2.22, w: 4.0, h: 0.05,
  fill: { color: C.orange },
  line: { color: C.orange },
});

// Key takeaways
const takeaways = [
  "CLIで動くAIアシスタント — インストールはnpmで簡単",
  "自然言語で指示するだけ — コーディング経験は不要",
  "スラッシュコマンドで操作 — まず /help から",
  "具体的な指示がカギ — ファイル名・エラーを明記",
  "CLAUDE.md でプロジェクト固有の文脈を共有",
];

takeaways.forEach((t, i) => {
  s8.addShape(pptx.ShapeType.rect, {
    x: 0.55, y: 2.48 + i * 0.7, w: 0.08, h: 0.4,
    fill: { color: C.orange },
    line: { color: C.orange },
  });
  s8.addText(t, {
    x: 0.82, y: 2.46 + i * 0.7, w: 8.0, h: 0.42,
    fontSize: 13.5,
    color: C.white,
    fontFace: "Calibri",
    valign: "middle",
  });
});

// Next step
s8.addShape(pptx.ShapeType.rect, {
  x: 0.55, y: 6.25, w: 8.5, h: 0.58,
  fill: { color: "026E7C" },
  line: { color: "026E7C" },
});
s8.addText("次のステップ:  claude と入力して、まず「Pythonでhello worldを出力するコードを書いて」と試してみよう！", {
  x: 0.75, y: 6.26, w: 8.1, h: 0.56,
  fontSize: 11.5,
  color: C.white,
  fontFace: "Calibri",
  bold: true,
  valign: "middle",
});

// ── Save ─────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: "usutaku_claudecode_lecture.pptx" })
  .then(() => console.log("✅ PPTX saved: usutaku_claudecode_lecture.pptx"))
  .catch(err => { console.error(err); process.exit(1); });
