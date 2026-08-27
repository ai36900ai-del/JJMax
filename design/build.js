const fs = require('fs');
const path = require('path');
const D = require('docx');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, VerticalAlign, ShadingType, HeightRule,
  TableLayoutType, Header, Footer, PageNumber, PageBreak, LevelFormat, FrameAnchorType,
} = D;

/* ---------- brand ---------- */
const NAVY   = '035891';   // sampled from cover artwork
const NAVYD  = '02385C';
const CYAN   = '18E5E9';   // sampled from cover artwork
const TEAL   = '0A7F92';
const SLATE  = '2E5F7F';   // business-model diagram
const BG     = 'F3F7FA';
const BGCY   = 'E9FBFC';
const LINE   = 'D6E0E8';
const TEXT   = '1A1D21';
const MUTED  = '5D6B78';
const WHITE  = 'FFFFFF';

const JP = 'Yu Gothic';
const F  = { ascii: 'Arial', eastAsia: JP, hAnsi: 'Arial', cs: 'Arial' };

const CW = 9638;                 // content width (A4, 20mm side margins)
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NOBORDER = { top: NONE, bottom: NONE, left: NONE, right: NONE,
                   insideHorizontal: NONE, insideVertical: NONE };
const line = (c = LINE, sz = 4) => ({ style: BorderStyle.SINGLE, size: sz, color: c });

/* ---------- run helpers ---------- */
const r = (text, o = {}) => new TextRun({ text, font: F, ...o });
const en = (text, o = {}) => new TextRun({ text, font: F, characterSpacing: 24, ...o });

/* ---------- paragraph helpers ---------- */
function body(text, o = {}) {
  return new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: { line: 300, after: 140, ...(o.spacing || {}) },
    children: [r(text, { size: 21, color: TEXT })],
  });
}
function lead(text) {
  return new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: { line: 320, after: 260 },
    children: [r(text, { size: 22, color: NAVYD })],
  });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'mars-bullet', level: 0 },
    alignment: AlignmentType.BOTH,
    spacing: { line: 290, after: 80 },
    children: [r(text, { size: 20, color: TEXT })],
  });
}
function h2(text, sub) {
  const kids = [r(text, { size: 24, bold: true, color: NAVYD })];
  if (sub) kids.push(r('   '), en(sub, { size: 15, color: MUTED }));
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: BG },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: NAVY, space: 8 } },
    indent: { left: 140, right: 100 },
    spacing: { before: 360, after: 220, line: 300 },
    children: kids,
  });
}
function spacer(h = 160) {
  return new Paragraph({ spacing: { after: h, line: 20 }, children: [r('', { size: 2 })] });
}
function pageBreak() {
  return new Paragraph({ spacing: { after: 0, line: 20 }, children: [new PageBreak()] });
}

/* ---------- table helpers ---------- */
function cell(children, o = {}) {
  return new TableCell({
    width: { size: o.w || CW, type: WidthType.DXA },
    columnSpan: o.span,
    shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
    margins: o.margins || { top: 120, bottom: 120, left: 170, right: 170 },
    verticalAlign: o.va || VerticalAlign.CENTER,
    borders: o.borders || NOBORDER,
    children,
  });
}
function table(rows, widths, o = {}) {
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: o.borders || NOBORDER,
    rows,
  });
}

/* ---------- chapter band ---------- */
function chapterBand(num, jp, enTitle) {
  const numCell = cell(
    [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: 520 },
      children: [r(num, { size: 56, bold: true, color: CYAN })],
    })],
    { w: 1500, fill: NAVYD, margins: { top: 200, bottom: 200, left: 80, right: 80 },
      borders: { ...NOBORDER, bottom: line(CYAN, 12) } },
  );
  const txtCell = cell(
    [
      new Paragraph({
        spacing: { before: 0, after: 60, line: 300 },
        children: [r(jp, { size: 30, bold: true, color: WHITE })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 0, line: 220 },
        children: [en(enTitle.toUpperCase(), { size: 15, color: CYAN })],
      }),
    ],
    { w: CW - 1500, fill: NAVY, margins: { top: 200, bottom: 200, left: 280, right: 200 },
      borders: { ...NOBORDER, bottom: line(CYAN, 12) } },
  );
  return table([new TableRow({
    height: { value: 1200, rule: HeightRule.ATLEAST },
    children: [numCell, txtCell],
  })], [1500, CW - 1500]);
}

/* ---------- numbered point card ---------- */
function pointCard(num, title, paras) {
  const nCell = cell(
    [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0, line: 240 },
        children: [en('POINT', { size: 11, color: CYAN })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0, line: 400 },
        children: [r(num, { size: 40, bold: true, color: WHITE })] }),
    ],
    { w: 1150, fill: SLATE, va: VerticalAlign.CENTER,
      margins: { top: 160, bottom: 160, left: 60, right: 60 } },
  );
  const kids = [new Paragraph({
    spacing: { before: 0, after: 120, line: 280 },
    children: [r(title, { size: 22, bold: true, color: NAVYD })],
  })];
  paras.forEach((p) => {
    if (Array.isArray(p)) {
      kids.push(new Paragraph({
        alignment: AlignmentType.BOTH,
        spacing: { after: 70, line: 280 },
        children: [r(p[0] + '　', { size: 19, bold: true, color: TEAL }),
                   r(p[1], { size: 19, color: TEXT })],
      }));
    } else {
      kids.push(new Paragraph({
        alignment: AlignmentType.BOTH,
        spacing: { after: 70, line: 280 },
        children: [r(p, { size: 19, color: TEXT })],
      }));
    }
  });
  const tCell = cell(kids, {
    w: CW - 1150, fill: BG, va: VerticalAlign.TOP,
    margins: { top: 180, bottom: 180, left: 240, right: 220 },
  });
  return [
    table([new TableRow({ children: [nCell, tCell] })], [1150, CW - 1150]),
    spacer(160),
  ];
}

/* ---------- key-value table ---------- */
function kvTable(rows) {
  const trs = rows.map(([k, vLines], i) => new TableRow({
    children: [
      cell([new Paragraph({ spacing: { after: 0, line: 280 },
        children: [r(k, { size: 20, bold: true, color: WHITE })] })],
        { w: 2600, fill: i % 2 ? NAVY : NAVYD, va: VerticalAlign.CENTER,
          margins: { top: 150, bottom: 150, left: 220, right: 160 },
          borders: { ...NOBORDER, bottom: line(WHITE, 6) } }),
      cell(vLines.map((v, j) => new Paragraph({
        spacing: { after: j === vLines.length - 1 ? 0 : 40, line: 280 },
        children: [r(v, { size: 20, color: TEXT })],
      })),
        { w: CW - 2600, fill: i % 2 ? 'FFFFFF' : BG, va: VerticalAlign.CENTER,
          margins: { top: 150, bottom: 150, left: 240, right: 200 },
          borders: { ...NOBORDER, bottom: line(LINE, 6) } }),
    ],
  }));
  return table(trs, [2600, CW - 2600]);
}

/* ---------- chip grid ---------- */
function chipGrid(items, cols = 3) {
  const w = Math.floor(CW / cols);
  const widths = Array(cols).fill(w);
  const rows = [];
  for (let i = 0; i < items.length; i += cols) {
    const slice = items.slice(i, i + cols);
    while (slice.length < cols) slice.push('');
    rows.push(new TableRow({
      height: { value: 620, rule: HeightRule.ATLEAST },
      children: slice.map((t) => cell(
        [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0, line: 260 },
          children: [r(t, { size: 19, bold: true, color: t ? NAVYD : 'FFFFFF' })] })],
        { w, fill: t ? BGCY : 'FFFFFF', va: VerticalAlign.CENTER,
          margins: { top: 130, bottom: 130, left: 100, right: 100 },
          borders: t ? { top: line(WHITE, 12), bottom: line(WHITE, 12),
                         left: line(WHITE, 12), right: line(WHITE, 12) } : NOBORDER },
      )),
    }));
  }
  return table(rows, widths);
}

/* ---------- callout ---------- */
function callout(label, text) {
  return table([new TableRow({
    children: [cell([
      new Paragraph({ spacing: { after: 80, line: 240 },
        children: [en(label.toUpperCase(), { size: 14, bold: true, color: TEAL })] }),
      new Paragraph({ alignment: AlignmentType.BOTH, spacing: { after: 0, line: 290 },
        children: [r(text, { size: 20, color: NAVYD })] }),
    ], { w: CW, fill: BGCY, va: VerticalAlign.TOP,
         margins: { top: 200, bottom: 200, left: 260, right: 240 },
         borders: { ...NOBORDER, left: { style: BorderStyle.SINGLE, size: 24, color: CYAN } } })],
  })], [CW]);
}

/* ---------- two-column comparison ---------- */
function twoCol(colA, colB) {
  const w = Math.floor(CW / 2);
  const head = (t, s, fill) => cell([
    new Paragraph({ spacing: { after: 40, line: 240 },
      children: [en(s.toUpperCase(), { size: 13, color: CYAN })] }),
    new Paragraph({ spacing: { after: 0, line: 280 },
      children: [r(t, { size: 22, bold: true, color: WHITE })] }),
  ], { w, fill, margins: { top: 170, bottom: 170, left: 220, right: 180 } });
  const bodyCell = (lines, fill) => cell(
    lines.map((l, i) => new Paragraph({
      alignment: AlignmentType.BOTH,
      spacing: { after: i === lines.length - 1 ? 0 : 90, line: 280 },
      children: Array.isArray(l)
        ? [r(l[0] + '　', { size: 19, bold: true, color: TEAL }), r(l[1], { size: 19, color: TEXT })]
        : [r(l, { size: 19, color: TEXT })],
    })),
    { w, fill, va: VerticalAlign.TOP, margins: { top: 190, bottom: 190, left: 220, right: 180 } },
  );
  return table([
    new TableRow({ children: [head(colA.title, colA.sub, NAVYD), head(colB.title, colB.sub, NAVY)] }),
    new TableRow({ children: [bodyCell(colA.lines, BG), bodyCell(colB.lines, 'FFFFFF')] }),
  ], [w, CW - w], { borders: { ...NOBORDER, insideVertical: line(WHITE, 12) } });
}

/* ---------- figure ---------- */
function figure(file, w, h, cap, num) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 140 },
      children: [new ImageRun({ data: fs.readFileSync(file), type: 'png',
        transformation: { width: w, height: h } })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200, line: 240 },
      border: { top: line(LINE, 4) },
      children: [en(num + '　', { size: 14, bold: true, color: TEAL }),
                 r(cap, { size: 17, color: MUTED })],
    }),
  ];
}

/* ---------- contents row ---------- */
function tocRow(num, jp, enTitle, page) {
  return new TableRow({
    height: { value: 760, rule: HeightRule.ATLEAST },
    children: [
      cell([new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 0, line: 400 },
        children: [r(num, { size: 36, bold: true, color: NAVY })] })],
        { w: 1100, va: VerticalAlign.CENTER, margins: { top: 160, bottom: 160, left: 60, right: 60 },
          borders: { ...NOBORDER, bottom: line(LINE, 6) } }),
      cell([
        new Paragraph({ spacing: { after: 30, line: 280 },
          children: [r(jp, { size: 24, bold: true, color: NAVYD })] }),
        new Paragraph({ spacing: { after: 0, line: 220 },
          children: [en(enTitle.toUpperCase(), { size: 13, color: MUTED })] }),
      ], { w: CW - 1100 - 1100, va: VerticalAlign.CENTER,
           margins: { top: 160, bottom: 160, left: 40, right: 60 },
           borders: { ...NOBORDER, bottom: line(LINE, 6) } }),
      cell([new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0, line: 280 },
        children: [en('P.' + page, { size: 19, bold: true, color: NAVY })] })],
        { w: 1100, va: VerticalAlign.CENTER, margins: { top: 160, bottom: 160, left: 40, right: 60 },
          borders: { ...NOBORDER, bottom: line(LINE, 6) } }),
    ],
  });
}

/* ================= document content ================= */
const M = (f) => path.join(__dirname, 'assets', f);

/* --- cover --- */
const cover = [
  new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [new ImageRun({
      data: fs.readFileSync(M('cover.png')), type: 'png',
      transformation: { width: 794, height: 1123 },
      floating: {
        horizontalPosition: { relative: D.HorizontalPositionRelativeFrom.PAGE, offset: 0 },
        verticalPosition: { relative: D.VerticalPositionRelativeFrom.PAGE, offset: 0 },
        behindDocument: true,
        allowOverlap: true,
        wrap: { type: D.TextWrappingType.NONE },
      },
    })],
  }),
  new Paragraph({
    spacing: { before: 12180, after: 100, line: 380 },
    indent: { left: 5400 },
    children: [r('株式会社 Mars Online', { size: 28, bold: true, color: WHITE })],
  }),
  new Paragraph({
    spacing: { before: 0, after: 0, line: 260 },
    indent: { left: 5400 },
    children: [en('ADVERTISING ／ CONSULTING ／ PLANNING', { size: 14, color: CYAN })],
  }),
];

/* --- contents --- */
const contents = [
  new Paragraph({ spacing: { before: 200, after: 40, line: 240 },
    children: [en('CONTENTS', { size: 16, bold: true, color: TEAL })] }),
  new Paragraph({
    spacing: { after: 80, line: 460 },
    children: [r('目　次', { size: 40, bold: true, color: NAVYD })],
  }),
  new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: CYAN } },
    spacing: { after: 420, line: 20 }, children: [r('', { size: 2 })],
  }),
  table([
    tocRow('01', '会社概要', 'Company Information', '03'),
    tocRow('02', '業界動向', 'Current Situation of the Industry', '04'),
    tocRow('03', '業界展望', 'Industry Prospects', '05'),
    tocRow('04', '事業展開計画', 'Company Development Plan', '06'),
  ], [1100, CW - 2200, 1100]),
  spacer(560),
  callout('Purpose of this document',
    '本書は、株式会社Mars Onlineが日本市場においてGoogle広告事業を展開するにあたり、'
    + '会社概要・業界動向・業界展望・事業展開計画を整理した事業計画書です。'),
];

/* --- 01 会社概要 --- */
const ch1 = [
  chapterBand('01', '会社概要', 'Company Information'),
  spacer(320),
  lead('株式会社Mars Onlineは、Google広告事業を主軸に、広告・コンサルティング・企画を'
     + '手がける東京・恵比寿発の企業です。'),
  h2('会社基本情報', 'Corporate Profile'),
  kvTable([
    ['会社名', ['株式会社Mars Online']],
    ['設立年月日', ['2025年2月26日']],
    ['本店所在地', ['東京都渋谷区恵比寿4丁目20番3号', '恵比寿ガーデンプレイスタワー18F']],
    ['主な事業内容', ['広告 ／ コンサルティング ／ 企画']],
  ]),
  spacer(400),
  h2('事業ドメイン', 'Business Domain'),
  chipGrid(['広告運用', 'コンサルティング', '企画・クリエイティブ'], 3),
  spacer(400),
  callout('Core business',
    '当社はGoogle広告事業を主軸としており、検索広告・ディスプレイ広告・動画広告・'
    + 'アプリ内広告を横断した広告運用と、その最適化を支える戦略設計を提供します。'),
];

/* --- 02 業界動向 --- */
const ch2 = [
  pageBreak(),
  chapterBand('02', '業界動向', 'Current Situation of the Industry'),
  spacer(320),
  lead('株式会社Mars OnlineはGoogle広告事業を主軸としており、Google広告の業界展望は'
     + '全体として良好である。主に以下の5つの側面に表れている。'),
  h2('現状を形づくる3つの要因', 'Three Driving Factors'),
  ...pointCard('1', 'Googleデジタル広告の市場規模の拡大', [
    '世界のデジタル広告支出は年々増加しており、特にモバイル広告と動画広告が急速に成長している。',
    ['広告主の予算シフト', '従来型メディアからデジタルチャネルへ予算を移す広告主が増えており、'
      + '世界最大級のデジタル広告プラットフォームであるGoogleは今後も継続的に恩恵を受ける。'],
  ]),
  ...pointCard('2', 'Google広告の技術的優位性', [
    ['精緻なターゲティング', 'Googleは強力なデータ分析力とAI技術を有しており、'
      + '精度の高い広告配信を実現し、広告効果を高めることができる。'],
    ['自動化ツール', 'スマート入札や広告最適化などの自動化ツールにより、'
      + '広告主は業務効率を高め、運用コストを低減できる。'],
  ]),
  ...pointCard('3', '多様な広告フォーマット', [
    '検索・ディスプレイ・動画・ショッピングの4フォーマットが、'
    + '目的の異なる広告主の需要をそれぞれ吸収している。',
  ]),
  spacer(120),
  table([
    new TableRow({
      children: [
        cell([new Paragraph({ spacing: { after: 0, line: 240 },
          children: [r('広告フォーマット', { size: 19, bold: true, color: WHITE })] })],
          { w: 2800, fill: NAVYD, margins: { top: 140, bottom: 140, left: 200, right: 140 } }),
        cell([new Paragraph({ spacing: { after: 0, line: 240 },
          children: [r('位置づけと需要', { size: 19, bold: true, color: WHITE })] })],
          { w: CW - 2800, fill: NAVY, margins: { top: 140, bottom: 140, left: 220, right: 180 } }),
      ],
    }),
    ...[
      ['検索広告', '中核事業として、検索広告は依然として重要な位置を占めている。'],
      ['ディスプレイ広告', 'Google Display Networkを通じて、広告主は数百万規模のウェブサイトやアプリ上に広告を掲載できる。'],
      ['動画広告', '世界第2位の検索エンジンであるYouTubeにおいて、動画広告の需要は伸び続けている。'],
      ['ショッピング広告', 'EC広告の需要は旺盛であり、ショッピング広告は小売事業者の商品露出と売上向上に寄与する。'],
    ].map(([k, v], i) => new TableRow({
      children: [
        cell([new Paragraph({ spacing: { after: 0, line: 260 },
          children: [r(k, { size: 19, bold: true, color: NAVYD })] })],
          { w: 2800, fill: i % 2 ? 'FFFFFF' : BG, va: VerticalAlign.TOP,
            margins: { top: 150, bottom: 150, left: 200, right: 140 },
            borders: { ...NOBORDER, bottom: line(LINE, 6) } }),
        cell([new Paragraph({ alignment: AlignmentType.BOTH, spacing: { after: 0, line: 280 },
          children: [r(v, { size: 19, color: TEXT })] })],
          { w: CW - 2800, fill: i % 2 ? 'FFFFFF' : BG, va: VerticalAlign.TOP,
            margins: { top: 150, bottom: 150, left: 220, right: 180 },
            borders: { ...NOBORDER, bottom: line(LINE, 6) } }),
      ],
    })),
  ], [2800, CW - 2800]),
];

/* --- 03 業界展望 --- */
const ch3 = [
  pageBreak(),
  chapterBand('03', '業界展望', 'Industry Prospects'),
  spacer(320),
  lead('モバイルシフトとグローバル×ローカライズの融合が、今後のGoogle広告市場の'
     + '成長余地を押し広げていく。'),
  h2('今後の成長を牽引する2つの潮流', 'Two Growth Drivers'),
  ...pointCard('4', 'モバイル広告の台頭', [
    ['モバイル端末の普及', 'スマートフォンやタブレットの広範な普及がモバイル広告の成長を牽引している。'],
    ['ローカライズ広告', 'モバイル端末はローカル広告の機会を大きく広げており、'
      + 'Googleはローカル検索広告やマップ広告によってこの需要に応えている。'],
  ]),
  ...pointCard('5', 'グローバル展開とローカライズの融合', [
    ['グローバルカバレッジ', 'Google広告プラットフォームは世界全域をカバーしており、'
      + '広告主が世界中のユーザーにリーチすることを可能にする。'],
    ['ローカライズサービス', 'Googleは各地域市場のニーズに応じたローカライズ広告サービスを提供している。'],
  ]),
  spacer(240),
  callout('Outlook',
    '以上の5つの側面から、Google広告の業界展望は全体として良好であると判断する。'
    + '当社はこの潮流の中で、日本市場に特化した運用力を強みに事業を展開する。'),
];

/* --- 04 事業展開計画 --- */
const ch4 = [
  pageBreak(),
  chapterBand('04', '事業展開計画', 'Company Development Plan'),
  spacer(320),
  lead('日本でGoogle広告事業を展開することには、以下の9つの顕著な優位性がある。'
     + '当社は日本の地の利を最大限に活用して事業を展開する。'),
  h2('日本市場を選ぶ理由', 'Why Japan'),
  chipGrid([
    '国際的な市場環境', '高度にデジタル化された市場', '旺盛な購買力',
    '優れた地理的立地', '先進的なテクノロジーインフラ', '良好な政策・法制度環境',
    'ローカライズされたサービス', '多様な広告フォーマット', 'データドリブンな精緻マーケティング',
  ], 3),
  spacer(360),
  h2('計画の進め方', 'Roadmap'),
  body('以下の3つの側面から計画を進める。'),
  spacer(120),
  ...pointCard('1', '市場分析とポジショニング', [
    ['市場調査', '日本のデジタル広告市場の現状・トレンド・競争環境を分析する。'],
    ['ターゲット層分析', '年齢・性別・所得水準・興味関心などの特性を明確化する。'],
    ['同業他社分析', '主要競合の広告戦略と市場シェアを研究したうえで、差別化の機会を見出す。'],
  ]),
  ...pointCard('2', '事業目標の明確化', [
    '短期目標と長期目標を設定し、段階的に日本市場での地位を確立する。',
  ]),
  twoCol(
    { title: '短期目標', sub: 'Within 6 months',
      lines: ['半年以内に一定数の顧客を獲得し、初期段階の市場浸透を実現する。'] },
    { title: '中長期目標', sub: 'Within 2 years',
      lines: ['2年以内に日本地域における主導的なGoogle広告サービスプロバイダーとなる。'] },
  ),
  pageBreak(),
  ...pointCard('3', '優位性のある顧客サービスの提供', [
    '全ラインナップの広告サービスと付加価値サービスを組み合わせ、'
    + '顧客のニーズに応じたカスタマイズ型の広告配信戦略・最適化プランを提示する。',
  ]),
  twoCol(
    { title: '広告サービス', sub: 'Advertising',
      lines: ['Google検索広告', 'ディスプレイ広告', '動画広告', 'アプリ内広告'] },
    { title: '付加価値サービス', sub: 'Value-added',
      lines: ['広告戦略コンサルティング', 'データ分析と最適化', '競合分析', '広告クリエイティブ制作'] },
  ),
];

/* --- business model --- */
const model = [
  spacer(360),
  h2('当社のビジネスモデル', 'Business Model'),
  body('開発者を中心に、Googleとの連携によってトラフィックの獲得と収益化を循環させ、'
     + 'インフラと現地専門チームがその基盤を支える構造とする。'),
  ...figure(M('model.png'), 540, 350, '開発者を中心としたトラフィック循環モデル', 'FIG.01'),
  spacer(220),
  callout('Structure',
    'トラフィック獲得はプラットフォーム買付＋共同運営提携により、収益化はGoogle広告により行う。'
    + 'その基盤として、クラウドサービス等の基盤インフラ、Google Cloud Partner＋'
    + '海外展開デベロッパーを支援する現地専門チームを置く。'),
];

/* --- back cover --- */
const back = [
  table([new TableRow({
    height: { value: 16838, rule: HeightRule.EXACT },
    children: [cell([
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120, line: 240 },
        children: [en('BUSINESS PLAN 2025', { size: 16, color: CYAN })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300, line: 460 },
        children: [r('株式会社Mars Online', { size: 36, bold: true, color: WHITE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300, line: 20 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: CYAN } },
        indent: { left: 2900, right: 2900 },
        children: [r('', { size: 2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60, line: 280 },
        children: [r('東京都渋谷区恵比寿4丁目20番3号', { size: 20, color: 'D9E6F0' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240, line: 280 },
        children: [r('恵比寿ガーデンプレイスタワー18F', { size: 20, color: 'D9E6F0' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0, line: 260 },
        children: [en('ADVERTISING ／ CONSULTING ／ PLANNING', { size: 14, color: CYAN })] }),
    ], { w: 11906, fill: NAVYD, va: VerticalAlign.CENTER,
         margins: { top: 0, bottom: 0, left: 800, right: 800 } })],
  })], [11906]),
];

/* ================= headers / footers ================= */
const hdr = new Header({
  children: [table([new TableRow({
    children: [
      cell([new Paragraph({ spacing: { after: 0, line: 240 },
        children: [r('株式会社Mars Online', { size: 15, color: MUTED })] })],
        { w: 4819, margins: { top: 0, bottom: 60, left: 0, right: 0 },
          borders: { ...NOBORDER, bottom: line(LINE, 4) } }),
      cell([new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0, line: 240 },
        children: [en('BUSINESS PLAN 2025', { size: 14, color: MUTED })] })],
        { w: 4819, margins: { top: 0, bottom: 60, left: 0, right: 0 },
          borders: { ...NOBORDER, bottom: line(LINE, 4) } }),
    ],
  })], [4819, 4819])],
});
const ftr = new Footer({
  children: [table([new TableRow({
    children: [
      cell([new Paragraph({ spacing: { after: 0, line: 240 },
        children: [en('MARS ONLINE INC.', { size: 13, color: MUTED })] })],
        { w: 4819, margins: { top: 60, bottom: 0, left: 0, right: 0 },
          borders: { ...NOBORDER, top: line(LINE, 4) } }),
      cell([new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0, line: 240 },
        children: [new TextRun({ font: F, size: 17, bold: true, color: NAVY,
          children: [PageNumber.CURRENT] })] })],
        { w: 4819, margins: { top: 60, bottom: 0, left: 0, right: 0 },
          borders: { ...NOBORDER, top: line(LINE, 4) } }),
    ],
  })], [4819, 4819])],
});

/* ================= assemble ================= */
const PAGE = { width: 11906, height: 16838 };
const doc = new Document({
  creator: '株式会社Mars Online',
  title: '事業計画書 2025 ｜ 株式会社Mars Online',
  description: 'Business Plan 2025',
  numbering: {
    config: [{
      reference: 'mars-bullet',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '▪', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 340, hanging: 220 } },
                 run: { color: TEAL, font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: JP } } },
      }],
    }],
  },
  styles: {
    default: {
      document: { run: { font: F, size: 21, color: TEXT } },
    },
  },
  sections: [
    { properties: { page: { size: PAGE, margin: { top: 0, right: 0, bottom: 0, left: 0,
        header: 0, footer: 0, gutter: 0 } } }, children: cover },
    { properties: { page: { size: PAGE, margin: { top: 1418, right: 1134, bottom: 1276, left: 1134,
        header: 680, footer: 566, gutter: 0 }, pageNumbers: { start: 2 } } },
      headers: { default: hdr }, footers: { default: ftr },
      children: [...contents, pageBreak(), ...ch1, ...ch2, ...ch3, ...ch4, ...model] },
    { properties: { page: { size: PAGE, margin: { top: 0, right: 0, bottom: 0, left: 0,
        header: 0, footer: 0, gutter: 0 } } }, children: back },
  ],
});

Packer.toBuffer(doc).then((b) => {
  fs.writeFileSync(path.join(__dirname, '..', '事業計画書_MarsOnline_2025.docx'), b);
  console.log('written', b.length);
});
