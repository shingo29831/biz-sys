/**
 * Module: @role: 応募書類コピペ支援、Excel試算表動的切替、およびWord貼付用Canvas図解画像生成
 */

// テーマ切替
const toggleBtn = document.getElementById('theme-toggle');
const themeIcon = document.getElementById('theme-icon');
const themeText = document.getElementById('theme-text');
const htmlElement = document.documentElement;
const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
const savedTheme = localStorage.getItem('theme');

function applyTheme(theme) {
  htmlElement.setAttribute('data-theme', theme);
  localStorage.setItem('theme', theme);
  if (themeIcon) themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
  if (themeText) themeText.textContent = theme === 'dark' ? 'ライトモードに切替' : 'ダークモードに切替';
}

if (savedTheme === 'dark' || (!savedTheme && prefersDark)) {
  applyTheme('dark');
} else {
  applyTheme('light');
}

if (toggleBtn) {
  toggleBtn.addEventListener('click', () => {
    const current = htmlElement.getAttribute('data-theme');
    applyTheme(current === 'dark' ? 'light' : 'dark');
  });
}

// テキストコピー機能
function copyText(elementId, btn) {
  const target = document.getElementById(elementId);
  if (!target) return;
  copyRaw(target.innerText, btn);
}

function showCopiedFeedback(btn, message, duration) {
  if (!btn.dataset.origHtml) btn.dataset.origHtml = btn.innerHTML;
  if (btn._copyTimer) clearTimeout(btn._copyTimer);
  btn.innerHTML = message;
  btn.classList.add('copied');
  btn._copyTimer = setTimeout(() => {
    btn.innerHTML = btn.dataset.origHtml;
    btn.classList.remove('copied');
    btn._copyTimer = null;
  }, duration);
}

function copyRaw(text, btn) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => {
      showCopiedFeedback(btn, '✅ コピー完了!', 1800);
    }).catch(() => fallbackCopyText(text, btn));
  } else {
    fallbackCopyText(text, btn);
  }
}

function fallbackCopyText(text, btn) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand('copy');
    showCopiedFeedback(btn, '✅ コピー完了!', 1800);
  } finally {
    document.body.removeChild(ta);
  }
}

// Excel表データ（月次 / 年次切替）
const excelData = {
  monthly: {
    revTitle: '② 中央表：売上（セル E8〜K14・月次1ヶ月基準）単位：円',
    expTitle: '③ 右表：経費・利益（セル M8〜V27・月次1ヶ月基準）単位：円',
    rev: [
      { row: 8, name: 'ヘビー室利用料（防音・PC・常設BS・1室平均単価）', price: 1800, qty: 324, note: '10.8h×30日' },
      { row: 10, name: 'ライト室利用料（クッション個室・2室平均単価）', price: 900, qty: 648, note: '10.8h×2室×30日' },
      { row: 12, name: '衛生シリコンカバー買い切り販売・追加機器貸出', price: 500, qty: 150, note: '月150件' },
      { row: 14, name: '室内ドリンク・アメニティ物販（実質粗利分）', price: 150, qty: 100, note: '月100本' }
    ],
    exp: [
      { row: 8, name: '店舗家賃・共益費（駅近約14坪・3室）', price: 300000, qty: 1 },
      { row: 10, name: '電気・水道光熱費（PC・空調24h稼働：1,859kWh）', price: 75000, qty: 1 },
      { row: 12, name: '通信費（法人向け10Gbps専有光回線・プロバイダ）', price: 20000, qty: 1 },
      { row: 14, name: 'オンライン予約・事前決済システム手数料（売上の4.5%）', price: 56538, qty: 1 },
      { row: 16, name: '室内除菌・衛生消耗品補充費（スタッフ内製清掃）', price: 15000, qty: 1 },
      { row: 18, name: 'VR機器保守更新積立金（HMD・トラッカー摩耗更新）', price: 65000, qty: 1 },
      { row: 20, name: '施設所有管理者賠償責任保険・動産総合保険', price: 15000, qty: 1 },
      { row: 22, name: '機器管理・清掃専任スタッフ人件費（1名・社保交通費込）', price: 265000, qty: 1 }
    ]
  },
  annual: {
    revTitle: '② 中央表：売上（セル E8〜K14・年次12ヶ月換算基準）単位：円',
    expTitle: '③ 右表：経費・利益（セル M8〜V27・年次12ヶ月換算基準）単位：円',
    rev: [
      { row: 8, name: 'ヘビー室利用料（防音・PC・常設BS・1室年間）', price: 1800, qty: 3888, note: '324h×12ヶ月' },
      { row: 10, name: 'ライト室利用料（クッション個室・2室年間）', price: 900, qty: 7776, note: '648h×12ヶ月' },
      { row: 12, name: '衛生シリコンカバー買い切り販売・追加機器貸出', price: 500, qty: 1800, note: '年1,800件' },
      { row: 14, name: '室内ドリンク・アメニティ物販（実質粗利分）', price: 150, qty: 1200, note: '年1,200本' }
    ],
    exp: [
      { row: 8, name: '店舗家賃・共益費（駅近約14坪・12ヶ月）', price: 300000, qty: 12 },
      { row: 10, name: '電気・水道光熱費（PC・空調24h稼働・12ヶ月）', price: 75000, qty: 12 },
      { row: 12, name: '通信費（法人向け10Gbps専有光回線・12ヶ月）', price: 20000, qty: 12 },
      { row: 14, name: 'オンライン予約・事前決済システム手数料（4.5%×12ヶ月）', price: 56538, qty: 12 },
      { row: 16, name: '室内除菌・衛生消耗品補充費（内製清掃・12ヶ月）', price: 15000, qty: 12 },
      { row: 18, name: 'VR機器保守更新積立金（HMD・トラッカー更新・12ヶ月）', price: 65000, qty: 12 },
      { row: 20, name: '施設所有管理者賠償責任保険・動産総合保険（12ヶ月）', price: 15000, qty: 12 },
      { row: 22, name: '機器管理・清掃専任スタッフ人件費（1名・12ヶ月）', price: 265000, qty: 12 }
    ]
  }
};

function switchExcelMode(mode, btn) {
  const buttons = document.querySelectorAll('.tab-btn');
  buttons.forEach(b => b.classList.remove('active'));
  const targetBtn = btn || buttons[mode === 'annual' ? 1 : 0];
  if (targetBtn) targetBtn.classList.add('active');
  const d = excelData[mode];
  const revTitleEl = document.getElementById('rev-table-title');
  const expTitleEl = document.getElementById('exp-table-title');
  if (revTitleEl) revTitleEl.textContent = d.revTitle;
  if (expTitleEl) expTitleEl.textContent = d.expTitle;

  let revSum = 0;
  const revHtml = d.rev.map(r => {
    const sub = r.price * r.qty;
    revSum += sub;
    return `<tr>
      <td><code>${r.row}行</code></td>
      <td>${r.name} <button class="copy-btn sm" onclick="copyRaw('${r.name}', this)">📋 品名</button></td>
      <td class="num-cell">${r.price.toLocaleString()} <button class="copy-btn sm" onclick="copyRaw('${r.price}', this)">📋</button></td>
      <td class="num-cell">${r.qty.toLocaleString()} <small>(${r.note})</small> <button class="copy-btn sm" onclick="copyRaw('${r.qty}', this)">📋</button></td>
      <td class="num-cell"><strong>${sub.toLocaleString()}</strong></td>
    </tr>`;
  }).join('') + `<tr class="highlight-cell">
    <td><code>K27</code></td>
    <td colspan="3"><strong>売上合計（自動計算：<code>=SUM(K8:K14)</code>）</strong></td>
    <td class="num-cell"><strong>${revSum.toLocaleString()} 円</strong></td>
  </tr>`;
  const revBodyEl = document.getElementById('excel-rev-body');
  if (revBodyEl) revBodyEl.innerHTML = revHtml;

  let expSum = 0;
  const expHtml = d.exp.map(r => {
    const sub = r.price * r.qty;
    expSum += sub;
    return `<tr>
      <td><code>${r.row}行</code></td>
      <td>${r.name} <button class="copy-btn sm" onclick="copyRaw('${r.name}', this)">📋 項目</button></td>
      <td class="num-cell">${r.price.toLocaleString()} <button class="copy-btn sm" onclick="copyRaw('${r.price}', this)">📋</button></td>
      <td class="num-cell">${r.qty.toLocaleString()} <button class="copy-btn sm" onclick="copyRaw('${r.qty}', this)">📋</button></td>
      <td class="num-cell"><strong>${sub.toLocaleString()}</strong></td>
    </tr>`;
  }).join('') + `<tr class="highlight-cell">
    <td><code>S27 / V27</code></td>
    <td colspan="2"><strong>経費合計（<code>S27</code>）：${expSum.toLocaleString()} 円</strong></td>
    <td colspan="2"><strong>利益（<code>V27 = K27 - S27</code>）：+${(revSum - expSum).toLocaleString()} 円（利益率 35.4%）</strong></td>
  </tr>`;
  const expBodyEl = document.getElementById('excel-exp-body');
  if (expBodyEl) expBodyEl.innerHTML = expHtml;
}

// 角丸矩形描画ユーティリティ
function roundRect(ctx, x, y, w, h, r, fill, stroke, lw = 2) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
  if (fill) { ctx.fillStyle = fill; ctx.fill(); }
  if (stroke) { ctx.lineWidth = lw; ctx.strokeStyle = stroke; ctx.stroke(); }
}

// 【図1】ビジネススキーム図描画
function drawSchemeDiagram() {
  const c = document.getElementById('canvas-scheme');
  if (!c) return;
  const ctx = c.getContext('2d');
  const font = '"Hiragino Kaku Gothic ProN", "Meiryo", sans-serif';
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, c.height);

  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 24px ${font}`;
  ctx.textAlign = 'center';
  ctx.fillText('【図1】防音VR個室「V-Room」ビジネススキーム図（ヒト・モノ・カネ・情報の流れ）', 600, 42);

  const topBoxes = [
    {
      x: 40,
      title: '① 物件オーナー・通信キャリア',
      l1: '・駅近ビル（約14坪/45㎡・3室構成）',
      l2: '・法人10Gbps専有光回線＋Wi-Fi 6E',
      l3: '【支払】家賃30万円＋光熱通信9.5万円/月'
    },
    {
      x: 420,
      title: '② VR・PC機材 / 衛生品ベンダー',
      l1: '・ハイエンドPC・Quest 3・常設BS',
      l2: '・衛生シリコンカバー卸（原価100円/個）',
      l3: '【支払】初期機材100万円＋保守保険8.0万円/月'
    },
    {
      x: 800,
      title: '③ 専任管理スタッフ1名・決済基盤',
      l1: '・利用間15分での直接清掃・機器検品・保守',
      l2: '・初心者サポート・Web事前決済基盤',
      l3: '【支払】人件費26.5万＋消耗品1.5万＋決済5.7万'
    }
  ];

  topBoxes.forEach(b => {
    roundRect(ctx, b.x, 70, 360, 150, 10, '#f8fafc', '#cbd5e1', 2);
    ctx.fillStyle = '#0f172a';
    ctx.font = `bold 16px ${font}`;
    ctx.fillText(b.title, b.x + 180, 102, 340);
    ctx.fillStyle = '#334155';
    ctx.font = `13px ${font}`;
    ctx.fillText(b.l1, b.x + 180, 134, 340);
    ctx.fillText(b.l2, b.x + 180, 162, 340);
    ctx.fillStyle = '#2563eb';
    ctx.font = `bold 13px ${font}`;
    ctx.fillText(b.l3, b.x + 180, 192, 340);
  });

  roundRect(ctx, 40, 235, 1120, 36, 8, '#f1f5f9', '#94a3b8', 1.5);
  ctx.fillStyle = '#2563eb';
  ctx.font = `bold 14px ${font}`;
  ctx.fillText('▲ 固定費・人件費・決済手数料の支払（月81.2万円） ｜ ▼ 空間・高速回線・有人検品済みの清潔な機材供給', 600, 258, 1100);

  roundRect(ctx, 40, 285, 1120, 125, 12, '#eff6ff', '#2563eb', 2.5);
  ctx.fillStyle = '#1e3a8a';
  ctx.font = `bold 20px ${font}`;
  ctx.fillText('🚀 当事業：VR特化型レンタルスペース「V-Room」（1店舗3室・専任スタッフ1名有人管理モデル）', 600, 322, 1100);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 15px ${font}`;
  ctx.fillText('有人インターバル直接清掃・機器保守 × スマートロック自動解錠 × 防音クッション空間 × UV-C除菌保管庫', 600, 355, 1100);
  ctx.fillStyle = '#475569';
  ctx.font = `14px ${font}`;
  ctx.fillText('秋葉原・新宿等の駅近約14坪（45㎡） ｜ ヘビー室1室（PC・防音・フルトラ） ＋ ライト室2室（体験・持込）', 600, 385, 1100);

  roundRect(ctx, 40, 425, 1120, 36, 8, '#f1f5f9', '#94a3b8', 1.5);
  ctx.fillStyle = '#2563eb';
  ctx.font = `bold 14px ${font}`;
  ctx.fillText('▼ 安全・清潔なVR空間＆選べる機材・買い切りカバー提供 ｜ ▲ 時間貸し利用料＋物販売上（標準月商 125.6万円／営業利益 44.5万円）', 600, 448, 1100);

  const botBoxes = [
    {
      x: 40,
      w: 540,
      title: '🧑‍🎓 ライト層（平日昼〜夕方中心）',
      l1: '自前HMD持込層・VR未経験者・購入検討層・学生サークル',
      l2: '利用料：持込 600円/時 ｜ 手ぶら（Quest 3込）1,000円/時（加重平均 900円/時）',
      l3: 'お得なパック：3h 1,500円〜 ｜ 深夜8h 2,400円〜（実質300円/時〜）',
      l4: '物販：買い切りシリコンカバー 500円（次回持参可・原価100円）＋ ドリンク粗利'
    },
    {
      x: 620,
      w: 540,
      title: '💃 ヘビー層（夜間・深夜・休日中心）',
      l1: 'VRChat定例集会・VRダンサー・VTuber配信/収録・音楽イベント参加者',
      l2: '利用料：持込 1,200円/時 ｜ 手ぶら（フルトラ込）2,000円/時（加重平均 1,800円/時）',
      l3: 'お得なパック：3h 3,000円〜 ｜ 深夜8h 4,500円〜（実質562円/時〜）',
      l4: '追加オプション：フルトラッカー単発貸出（+500円/回）＋ 10Gbps有線防音環境'
    }
  ];

  botBoxes.forEach(b => {
    roundRect(ctx, b.x, 475, b.w, 165, 10, '#f8fafc', '#cbd5e1', 2);
    ctx.fillStyle = '#0f172a';
    ctx.font = `bold 16px ${font}`;
    ctx.fillText(b.title, b.x + b.w / 2, 506, b.w - 30);
    ctx.fillStyle = '#475569';
    ctx.font = `13px ${font}`;
    ctx.fillText(b.l1, b.x + b.w / 2, 534, b.w - 30);
    ctx.fillStyle = '#0f172a';
    ctx.font = `bold 13px ${font}`;
    ctx.fillText(b.l2, b.x + b.w / 2, 562, b.w - 30);
    ctx.fillStyle = '#2563eb';
    ctx.fillText(b.l3, b.x + b.w / 2, 588, b.w - 30);
    ctx.fillStyle = '#059669';
    ctx.font = `12.5px ${font}`;
    ctx.fillText(b.l4, b.x + b.w / 2, 614, b.w - 30);
  });

  roundRect(ctx, 40, 655, 1120, 48, 8, '#f0fdf4', '#16a34a', 2);
  ctx.fillStyle = '#15803d';
  ctx.font = `bold 15px ${font}`;
  ctx.fillText('★「ヘビー1室＋物販」だけで固定費の89.2%を回収し、ライト2室が利益化（稼働45%で営業利益+44.5万円・約9.0ヶ月回収）', 600, 685, 1100);

  const img = document.getElementById('img-scheme');
  if (img) img.src = c.toDataURL('image/png');
}

// 【図2】国内メタバース市場予測＆競合ポジショニング図描画
function drawMarketDiagram() {
  const c = document.getElementById('canvas-market');
  if (!c) return;
  const ctx = c.getContext('2d');
  const font = '"Hiragino Kaku Gothic ProN", "Meiryo", sans-serif';
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, c.height);

  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 22px ${font}`;
  ctx.textAlign = 'center';
  ctx.fillText('【図2】国内メタバース市場の急拡大予測 ＆ 競合ポジショニング比較マトリクス', 600, 38);

  roundRect(ctx, 30, 60, 555, 535, 12, '#f8fafc', '#cbd5e1', 2);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 17px ${font}`;
  ctx.fillText('① 国内メタバース市場予測（矢野経済研究所 2024年調査）', 307, 92, 530);

  const bars = [
    { year: '2023年度', text: '1,863億円', h: 32 },
    { year: '2024年度(見込)', text: '2,750億円', h: 48 },
    { year: '2026年度(予)', text: '7,500億円', h: 128 },
    { year: '2028年度(予)', text: '1兆8,700億円', h: 260 }
  ];
  const chartBaseY = 390;
  ctx.strokeStyle = '#94a3b8';
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(55, chartBaseY);
  ctx.lineTo(560, chartBaseY);
  ctx.stroke();

  bars.forEach((b, idx) => {
    const bx = 80 + idx * 122;
    const bw = 85;
    ctx.fillStyle = idx === 3 ? '#2563eb' : '#38bdf8';
    ctx.fillRect(bx, chartBaseY - b.h, bw, b.h);

    ctx.fillStyle = '#0f172a';
    ctx.font = `bold ${idx === 3 ? '13px' : '12px'} ${font}`;
    ctx.fillText(b.text, bx + bw / 2, chartBaseY - b.h - 8, bw + 30);

    ctx.fillStyle = '#475569';
    ctx.font = `11.5px ${font}`;
    ctx.fillText(b.year, bx + bw / 2, chartBaseY + 20, bw + 20);
  });

  // 2028年度バー(x=446..531, y=130..390)との重なりを防ぐため左上空白域へ配置
  roundRect(ctx, 68, 135, 230, 36, 6, '#eff6ff', '#2563eb', 1.5);
  ctx.fillStyle = '#1e40af';
  ctx.font = `bold 13px ${font}`; ctx.fillText('約6.8倍急拡大（年平均+61.5\%）', 183, 158, 215); >>>> <<<< const compData = [ { name: '当事業「V-Room」', space: '◎ 3〜10畳防音', equip: '◎ 持込/常設BS', clean: '◎ 買い切り+除菌', price: '600〜2,000円/h', hl: true }, { name: '自宅単身住宅', space: '× 家具衝突・騒音', equip: '△ 初期30〜40万', clean: '◎ 個人所有', price: '家賃負担', hl: false }, ==== const compData = [ { name: '当事業「V-Room」', space: '◎ 3〜10畳防音', equip: '◎ 持込/常設BS', clean: '◎ 有人清掃+除菌', price: '600〜2,000円/h', hl: true }, { name: '自宅単身住宅', space: '× 家具衝突・騒音', equip: '△ 初期30〜40万', clean: '◎ 個人所有', price: '家賃負担', hl: false }, >>>> <<<< ctx.fillStyle = '#059669'; ctx.font = `bold 12.5px ${font}`;
  ctx.fillText('★買い切りシリコンカバーによる衛生対策で共用VRの最大の壁を突破', 890, 548, 500);

  const img = document.getElementById('img-market');
====
  ctx.fillStyle = '#059669';
  ctx.font = `bold 12.5px ${font}`; ctx.fillText('★専任スタッフの利用間清掃＋買い切りカバー＋UV除菌で衛生の壁を突破', 890, 548, 500);  const img = document.getElementById('img-market'); >>>> <<<< // 棒2: ヘビー室1室＋物販 67.3万円 ctx.fillStyle = '#4f46e5'; ctx.fillRect(250, baseY - 134, 110, 134); ctx.fillStyle = '#10b981'; ctx.fillRect(250, baseY - 155, 110, 21); ctx.fillStyle = '#0f172a'; ctx.font = `bold 14px ${font}`;
  ctx.fillText('ヘビー1室＋物販', 305, baseY + 22);
====
  // 棒2: ヘビー室1室＋物販 67.3万円
  ctx.fillStyle = '#4f46e5';
  ctx.fillRect(250, baseY - 134, 110, 134);
  ctx.fillStyle = '#10b981';
  ctx.fillRect(250, baseY - 155, 110, 21);
  ctx.fillStyle = '#ffffff';
  ctx.font = `bold 12px ${font}`; ctx.fillText('ヘビー1室 58.3万', 305, baseY - 62); ctx.fillStyle = '#0f172a'; ctx.font = `bold 14px ${font}`;
  ctx.fillText('ヘビー1室＋物販', 305, baseY + 22);

  roundRect(ctx, 45, 430, 525, 145, 8, '#ffffff', '#cbd5e1', 1.5);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 13.5px ${font}`;
  ctx.fillText('【熱狂的な需要 vs 狭小住宅ギャップ】', 307, 456, 500);
  ctx.fillStyle = '#334155';
  ctx.font = `12.5px ${font}`;
  ctx.fillText('・VRChat国別アクセス比率：日本が世界1位（37%・Sensor Tower）', 307, 482, 500);
  ctx.fillText('・首都圏単身向け賃貸床面積：平均20.3㎡（総務省住宅統計）', 307, 506, 500);
  ctx.fillStyle = '#dc2626';
  ctx.font = `bold 12.5px ${font}`;
  ctx.fillText('⇒ 自宅では不可能な「安全な可動空間・防音・フルトラ」への強い需要が存在', 307, 532, 500);

  roundRect(ctx, 610, 60, 560, 535, 12, '#f8fafc', '#cbd5e1', 2);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 17px ${font}`;
  ctx.fillText('② 既存サービスとの差別化（市場の空白地帯）', 890, 92, 530);

  const compData = [
    { name: '当事業「V-Room」', space: '◎ 3〜10畳防音', equip: '◎ 持込/常設BS', clean: '◎ 買い切り+除菌', price: '600〜2,000円/h', hl: true },
    { name: '自宅単身住宅', space: '× 家具衝突・騒音', equip: '△ 初期30〜40万', clean: '◎ 個人所有', price: '家賃負担', hl: false },
    { name: 'ネットカフェ', space: '× 着座1畳のみ', equip: '× VR非対応', clean: '△ 共用皮脂懸念', price: '500〜700円/h', hl: false },
    { name: '商業VR施設', space: '○ プレイゾーン', equip: '× 指定ゲームのみ', clean: '△ 不織布のみ', price: '約3,000円/15分', hl: false },
    { name: '業務用スタジオ', space: '◎ 15畳以上', equip: '○ モーションキャプチャ', clean: '○ 専門清掃', price: '15,000〜3万円/h', hl: false }
  ];

  const tableTop = 115;
  const colX = [625, 745, 835, 930, 1040, 1155];
  ctx.fillStyle = '#e2e8f0';
  ctx.fillRect(colX[0], tableTop, colX[5] - colX[0], 32);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 12.5px ${font}`;
  ctx.fillText('分類', (colX[0] + colX[1]) / 2, tableTop + 21);
  ctx.fillText('可動空間', (colX[1] + colX[2]) / 2, tableTop + 21);
  ctx.fillText('VR機材', (colX[2] + colX[3]) / 2, tableTop + 21);
  ctx.fillText('衛生管理', (colX[3] + colX[4]) / 2, tableTop + 21);
  ctx.fillText('時間単価目安', (colX[4] + colX[5]) / 2, tableTop + 21);

  compData.forEach((row, i) => {
    const rowY = tableTop + 32 + i * 54;
    ctx.fillStyle = row.hl ? '#e0f2fe' : (i % 2 === 0 ? '#ffffff' : '#f8fafc');
    ctx.fillRect(colX[0], rowY, colX[5] - colX[0], 54);
    ctx.strokeStyle = row.hl ? '#0284c7' : '#cbd5e1';
    ctx.lineWidth = row.hl ? 2 : 1;
    ctx.strokeRect(colX[0], rowY, colX[5] - colX[0], 54);

    ctx.fillStyle = row.hl ? '#0369a1' : '#0f172a';
    ctx.font = `${row.hl ? 'bold ' : ''}12px ${font}`;
    ctx.fillText(row.name, (colX[0] + colX[1]) / 2, rowY + 31, colX[1] - colX[0] - 6);
    ctx.fillStyle = row.hl ? '#0f172a' : '#334155';
    ctx.font = `11.5px ${font}`;
    ctx.fillText(row.space, (colX[1] + colX[2]) / 2, rowY + 31, colX[2] - colX[1] - 4);
    ctx.fillText(row.equip, (colX[2] + colX[3]) / 2, rowY + 31, colX[3] - colX[2] - 4);
    ctx.fillText(row.clean, (colX[3] + colX[4]) / 2, rowY + 31, colX[4] - colX[3] - 4);
    ctx.fillStyle = row.hl ? '#0284c7' : '#0f172a';
    ctx.font = `${row.hl ? 'bold ' : ''}11.5px ${font}`;
    ctx.fillText(row.price, (colX[4] + colX[5]) / 2, rowY + 31, colX[5] - colX[4] - 6);
  });

  roundRect(ctx, 625, 435, 530, 140, 8, '#eff6ff', '#2563eb', 1.5);
  ctx.fillStyle = '#1e3a8a';
  ctx.font = `bold 13.5px ${font}`;
  ctx.fillText('【競合ポジショニングの核心：市場の空白地帯】', 890, 465, 500);
  ctx.fillStyle = '#0f172a';
  ctx.font = `12.5px ${font}`;
  ctx.fillText('「手頃な価格（実質300円/h〜）」で「全身動ける防音空間」と「フルトラ常設」を', 890, 495, 500);
  ctx.fillText('兼ね備えた施設は日本に存在せず、ブルーオーシャンを獲得できる。', 890, 520, 500);
  ctx.fillStyle = '#059669';
  ctx.font = `bold 12.5px ${font}`;
  ctx.fillText('★買い切りシリコンカバーによる衛生対策で共用VRの最大の壁を突破', 890, 548, 500);

  const img = document.getElementById('img-market');
  if (img) img.src = c.toDataURL('image/png');
}

// 【図3】ロードマップ図描画
function drawRoadmapDiagram() {
  const c = document.getElementById('canvas-roadmap');
  if (!c) return;
  const ctx = c.getContext('2d');
  const font = '"Hiragino Kaku Gothic ProN", "Meiryo", sans-serif';
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, c.height);

  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 22px ${font}`;
  ctx.textAlign = 'center';
  ctx.fillText('【図3】5年後・10年後の拠点展開＆グループ年商成長ロードマップ図', 600, 38);

  const steps = [
    {
      x: 35,
      w: 260,
      phase: '1年目（実証＆ドミナント）',
      metric: '直営 1〜3店 ｜ 年商 4,520万円',
      desc1: '・秋葉原1号店：回収約9.0ヶ月',
      desc2: '・新宿・池袋へ3店舗ドミナント',
      desc3: '・VRChat連動・学割で稼働45%確立',
      hl: false
    },
    {
      x: 325,
      w: 260,
      phase: '3年目（FC展開開始）',
      metric: '全国 15拠点 ｜ 年商 2.25億円',
      desc1: '・東名阪・政令指定都市へ展開',
      desc2: '・1名有人管理FCパッケージ供給',
      desc3: '・空きビル高収益転換モデル確立',
      hl: false
    },
    {
      x: 615,
      w: 260,
      phase: '5年後（全国都市圏網羅）',
      metric: '全国 50拠点 ｜ 年商 7.5億円',
      desc1: '・直営10店 / FC40店網羅体制',
      desc2: '・メタバース1.87兆円市場に乗る',
      desc3: '・VTuber事務所の公式収録拠点化',
      hl: true
    },
    {
      x: 905,
      w: 260,
      phase: '10年後（社会インフラ化）',
      metric: '全国 150拠点 ｜ 年商 22.5億円',
      desc1: '・地方中核都市駅前まで網羅',
      desc2: '・企業XR遠隔技術研修拠点（BtoB）',
      desc3: '・自治体教育/不登校支援（BtoG）',
      hl: true
    }
  ];

  steps.forEach((s, idx) => {
    roundRect(ctx, s.x, 65, s.w, 245, 10, s.hl ? '#e0f2fe' : '#f8fafc', s.hl ? '#0284c7' : '#cbd5e1', s.hl ? 2.5 : 1.5);

    ctx.fillStyle = s.hl ? '#0369a1' : '#0f172a';
    ctx.font = `bold 15px ${font}`;
    ctx.fillText(s.phase, s.x + s.w / 2, 96, s.w - 20);

    roundRect(ctx, s.x + 10, 112, s.w - 20, 32, 6, s.hl ? '#0284c7' : '#2563eb', null);
    ctx.fillStyle = '#ffffff';
    ctx.font = `bold 12.5px ${font}`;
    ctx.fillText(s.metric, s.x + s.w / 2, 133, s.w - 26);

    ctx.fillStyle = '#334155';
    ctx.font = `12.5px ${font}`;
    ctx.fillText(s.desc1, s.x + s.w / 2, 175, s.w - 20);
    ctx.fillText(s.desc2, s.x + s.w / 2, 205, s.w - 20);
    ctx.fillText(s.desc3, s.x + s.w / 2, 235, s.w - 20);

    if (idx < 3) {
      ctx.fillStyle = '#2563eb';
      ctx.font = `bold 22px ${font}`;
      ctx.fillText('▶', s.x + s.w + 15, 185);
    }
  });

  roundRect(ctx, 35, 335, 1130, 215, 12, '#f0fdf4', '#16a34a', 2.5);
  ctx.fillStyle = '#15803d';
  ctx.font = `bold 18px ${font}`;
  ctx.fillText('【事業の継続性（サステナビリティ）と長期安定収益の3大根拠】', 600, 372, 1090);

  ctx.fillStyle = '#0f172a';
  ctx.font = `14px ${font}`;
  ctx.fillText('① 低い損益分岐点（BEP稼働率 28.3%＝1日6.8h）：スマートロック自動受付で1名体制に省人化し清掃を内製化。稼働30%の悲観時でも黒字維持。', 600, 410, 1090);
  ctx.fillText('② 計画的な機材更新積立（月6.5万円/店）：年間78万円の保守更新積立金を経費計上し、追加融資なしで2〜3年周期の最新HMD・PCへ更新。', 600, 444, 1090);
  ctx.fillText('③ BtoC×BtoB×BtoGのハイブリッド稼働：夜間・休日のエンタメ需要に加え、平日昼間は企業XR研修・自治体教育利用で24時間稼働を最大化。', 600, 478, 1090);

  ctx.fillStyle = '#1e40af';
  ctx.font = `bold 14px ${font}`;
  ctx.fillText('国内メタバース市場1兆8,700億円（2028年度予測）の成長波に乗り、日本の狭小住宅課題を解決する社会インフラへ成長', 600, 516, 1090);

  const img = document.getElementById('img-roadmap');
  if (img) img.src = c.toDataURL('image/png');
}

// 【図4】採算性証明＆清掃プロトコル図描画
function drawProfitHygieneDiagram() {
  const c = document.getElementById('canvas-profit-hygiene');
  if (!c) return;
  const ctx = c.getContext('2d');
  const font = '"Hiragino Kaku Gothic ProN", "Meiryo", sans-serif';
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, c.height);

  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 22px ${font}`;
  ctx.textAlign = 'center';
  ctx.fillText('【図4】「ライト2室＋ヘビー1室」で高収益を生む二階建て採算構造 ＆ 利用間15分清掃システム', 600, 38);

  roundRect(ctx, 30, 60, 555, 535, 12, '#f8fafc', '#cbd5e1', 2);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 18px ${font}`;
  ctx.fillText('① なぜ3室で採算が取れるのか？（二階建て収益構造）', 307, 94, 530);

  const baseY = 420;
  ctx.strokeStyle = '#94a3b8';
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(60, baseY);
  ctx.lineTo(555, baseY);
  ctx.stroke();

  // 棒1: 月間固定費 75.5万円
  ctx.fillStyle = '#ef4444';
  ctx.fillRect(90, baseY - 174, 110, 174);
  ctx.fillStyle = '#dc2626';
  ctx.font = `bold 14px ${font}`;
  ctx.fillText('75.5万円', 145, baseY - 182);
  ctx.fillStyle = '#0f172a';
  ctx.fillText('月間固定費合計', 145, baseY + 22);
  ctx.font = `12px ${font}`;
  ctx.fillStyle = '#475569';
  ctx.fillText('(専任スタッフ1名人件費込)', 145, baseY + 40);

  // 棒2: ヘビー室1室＋物販 67.3万円
  ctx.fillStyle = '#4f46e5';
  ctx.fillRect(250, baseY - 134, 110, 134);
  ctx.fillStyle = '#10b981';
  ctx.fillRect(250, baseY - 155, 110, 21);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 14px ${font}`;
  ctx.fillText('ヘビー1室＋物販', 305, baseY + 22);
  ctx.font = `bold 12px ${font}`;
  ctx.fillStyle = '#dc2626';
  ctx.fillText('★固定費の89.2%を回収!', 305, baseY + 40);

  // 棒3: 3室合計＋物販 125.6万円
  ctx.fillStyle = '#4f46e5';
  ctx.fillRect(415, baseY - 134, 110, 134);
  ctx.fillStyle = '#0ea5e9';
  ctx.fillRect(415, baseY - 268, 110, 134);
  ctx.fillStyle = '#10b981';
  ctx.fillRect(415, baseY - 290, 110, 22);
  ctx.fillStyle = '#047857';
  ctx.font = `bold 14px ${font}`;
  ctx.fillText('総売上 125.6万円', 470, baseY - 298);
  ctx.fillStyle = '#ffffff';
  ctx.font = `bold 12px ${font}`;
  ctx.fillText('ライト2室 58.3万', 470, baseY - 195);
  ctx.fillText('ヘビー1室 58.3万', 470, baseY - 62);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 14px ${font}`;
  ctx.fillText('3室＋物販合計', 470, baseY + 22);
  ctx.font = `bold 12px ${font}`;
  ctx.fillStyle = '#047857';
  ctx.fillText('★営業利益 +44.5万円!', 470, baseY + 40);

  // 固定費破線
  ctx.strokeStyle = '#dc2626';
  ctx.setLineDash([6, 4]);
  ctx.beginPath();
  ctx.moveTo(65, baseY - 174);
  ctx.lineTo(550, baseY - 174);
  ctx.stroke();
  ctx.setLineDash([]);

  // 棒2ラベルピル
  roundRect(ctx, 260, baseY - 198, 90, 20, 4, '#ffffff', '#4338ca', 1.2);
  ctx.fillStyle = '#4338ca';
  ctx.font = `bold 13px ${font}`;
  ctx.fillText('67.3万円', 305, baseY - 183);

  // 左下サマリーボックス
  roundRect(ctx, 50, 475, 515, 102, 8, '#eff6ff', '#2563eb', 1.8);
  ctx.fillStyle = '#1e3a8a';
  ctx.font = `bold 14px ${font}`;
  ctx.fillText('・損益分岐点（BEP）：1室あたり1日約6.8時間（稼働率28.3%）で黒字化', 307, 500, 495);
  ctx.fillText('・スマートロック自動受付で1名体制に省人化し、日常清掃の内製化を実現', 307, 526, 495);
  ctx.fillText('・「ヘビー1＋ライト2室」構成により有人管理でも約9.0ヶ月で投資回収！', 307, 552, 495);

  // 右側：4段階清掃プロトコル
  roundRect(ctx, 605, 60, 565, 535, 12, '#f8fafc', '#cbd5e1', 2);
  ctx.fillStyle = '#0f172a';
  ctx.font = `bold 18px ${font}`;
  ctx.fillText('② 利用と利用の間の清掃（専任スタッフ1名による4段階管理）', 887, 94, 535);

  const hygieneSteps = [
    {
      y: 115,
      title: 'Step 1：予約間の「自動15分インターバル」確保',
      desc1: '・予約枠と次の予約枠の間にシステムで「15分間の入室不可時間」を自動挿入',
      desc2: '・前利用者の熱気残留や鉢合わせを防ぎ、スタッフの入室検品時間を確保',
      bg: '#e0f2fe',
      ec: '#0284c7'
    },
    {
      y: 225,
      title: 'Step 2：専任スタッフによる「直接入室検品・除菌清拭」',
      desc1: '・スタッフが毎回入室しHMDレンズ皮脂確認・床とクッション壁の除菌清拭を実施',
      desc2: '・ケーブル断線・ねじれ解消、トラッカー充電装着、SteamVR境界線ズレを補正',
      bg: '#f0fdf4',
      ec: '#16a34a'
    },
    {
      y: 335,
      title: 'Step 3：「UV-C＋オゾン殺菌庫」＆15分間の空気100%急速換気',
      desc1: '・スマート保管庫内で波長254nmのUV-C紫外線照射＋オゾン脱臭を10分間実行',
      desc2: '・24h機械換気＋業務用プラズマクラスター脱臭機で15分間に室内空気を完全置換',
      bg: '#fefce8',
      ec: '#ca8a04'
    },
    {
      y: 445,
      title: 'Step 4：買い切りカバー必須化 ＆ 初心者・トラブル即時サポート',
      desc1: '・買い切りシリコンカバー(500円)＋使い捨てマスクでHMD本体への皮脂付着ゼロ',
      desc2: '・VR未経験者の初期装着サポートやペアリング不良にも専任スタッフが即応',
      bg: '#eef2ff',
      ec: '#4f46e5'
    }
  ];

  hygieneSteps.forEach((s, i) => {
    roundRect(ctx, 625, s.y, 525, 92, 10, s.bg, s.ec, 2);
    ctx.fillStyle = '#0f172a';
    ctx.font = `bold 15px ${font}`;
    ctx.fillText(s.title, 887, s.y + 26, 500);
    ctx.fillStyle = '#334155';
    ctx.font = `13.5px ${font}`;
    ctx.fillText(s.desc1, 887, s.y + 52, 500);
    ctx.fillText(s.desc2, 887, s.y + 74, 500);
    if (i < 3) {
      ctx.fillStyle = '#2563eb';
      ctx.font = `bold 16px ${font}`;
      ctx.fillText('▼', 887, s.y + 108);
    }
  });

  ctx.fillStyle = '#047857';
  ctx.font = `bold 13.5px ${font}`;
  ctx.fillText('★専任スタッフの直接管理により「清潔度100%・機器トラブルによる返金ゼロ」を実現', 887, 568, 535);

  const img = document.getElementById('img-profit-hygiene');
  if (img) img.src = c.toDataURL('image/png');
}

// Canvas画像をクリップボードにPNGコピー（Word貼付用）
function copyCanvasImage(canvasId, btn) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  canvas.toBlob(blob => {
    if (!blob) return;
    if (navigator.clipboard && window.ClipboardItem) {
      const item = new ClipboardItem({ 'image/png': blob });
      navigator.clipboard.write([item]).then(() => {
        showCopiedFeedback(btn, '✅ 画像をコピーしました！（WordでCtrl+V貼付可）', 2500);
      }).catch(() => {
        alert('ブラウザの制限により直接コピーできない場合は、下の画像を右クリックして「画像をコピー」を選択してください。');
      });
    } else {
      alert('下の画像を右クリックして「画像をコピー」を選択してください。');
    }
  }, 'image/png');
}

// Canvas画像のダウンロード
function downloadCanvasImage(canvasId, filename) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const link = document.createElement('a');
  link.download = filename;
  link.href = canvas.toDataURL('image/png');
  link.click();
}

// 初期実行
window.addEventListener('DOMContentLoaded', () => {
  switchExcelMode('monthly');
  drawSchemeDiagram();
  drawMarketDiagram();
  drawRoadmapDiagram();
  drawProfitHygieneDiagram();
});
