/**
 * Module: @role: プレゼンテーション制御（キーボードショートカット操作・全画面切替・Canvas図表描画・PDF/PPTXダウンロード生成）
 */
(function () {
  'use strict';

  const STAGE_WIDTH = 1600;
  const STAGE_HEIGHT = 900;

  const slides = Array.from(document.querySelectorAll('.slide'));
  const stage = document.getElementById('deck-stage');
  const overlay = document.getElementById('shortcut-overlay');
  const toast = document.getElementById('export-toast');

  let currentIndex = 0;
  let isExporting = false;
  let toastTimer = null;

  /* =========================================================
     1. ステージの16:9自動スケーリングとスライド遷移
  ========================================================= */
  function fitStageToWindow() {
    if (!stage) return;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const scale = Math.min(vw / STAGE_WIDTH, vh / STAGE_HEIGHT);
    stage.style.transform = `scale(${scale})`;
  }

  function showSlide(index) {
    if (index < 0 || index >= slides.length) return;
    closePreDeckScreen();
    currentIndex = index;
    slides.forEach((s, i) => {
      s.classList.toggle('active', i === currentIndex);
    });
    window.history.replaceState(null, '', `#slide-${currentIndex + 1}`);
  }

  function nextSlide() {
    if (currentIndex < slides.length - 1) {
      showSlide(currentIndex + 1);
    }
  }

  function prevSlide() {
    if (currentIndex > 0) {
      showSlide(currentIndex - 1);
    }
  }

  function showToast(message, durationMs = 3000) {
    if (!toast) return;
    if (toastTimer) clearTimeout(toastTimer);
    toast.textContent = message;
    toast.classList.remove('hidden');
    if (durationMs > 0) {
      toastTimer = setTimeout(() => {
        toast.classList.add('hidden');
      }, durationMs);
    }
  }

  /* =========================================================
     2. 全画面表示（Fullscreen API）切替
  ========================================================= */
  function toggleFullscreen() {
    closePreDeckScreen();
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(() => {
        showToast('全画面表示がブラウザによりブロックされました', 2500);
      });
    } else {
      document.exitFullscreen();
    }
  }

  function toggleShortcutHelp() {
    if (!overlay) return;
    const isHidden = overlay.classList.toggle('hidden');
    overlay.setAttribute('aria-hidden', String(isHidden));
  }

  /* =========================================================
     3. PDFダウンロード生成（P キー：全10スライドを16:9 PDF化）
  ========================================================= */
  async function exportToPDF() {
    if (isExporting) return;
    closePreDeckScreen();

    if (typeof window.html2canvas === 'undefined' || typeof window.jspdf === 'undefined') {
      window.print();
      return;
    }

    isExporting = true;
    const originalIndex = currentIndex;
    const prevTransform = stage.style.transform;

    try {
      showToast(`📄 PDFファイルを生成中 (1 / ${slides.length})...`, 0);
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'px',
        format: [STAGE_WIDTH, STAGE_HEIGHT],
        compress: true
      });

      stage.style.transform = 'none';

      for (let i = 0; i < slides.length; i++) {
        showToast(`📄 PDFファイルを生成中 (${i + 1} / ${slides.length})...`, 0);
        slides.forEach((s, idx) => s.classList.toggle('active', idx === i));

        await new Promise((r) => setTimeout(r, 60));

        const canvas = await window.html2canvas(slides[i], {
          scale: 1.5,
          width: STAGE_WIDTH,
          height: STAGE_HEIGHT,
          backgroundColor: '#ffffff',
          useCORS: true,
          logging: false
        });

        const imgData = canvas.toDataURL('image/jpeg', 0.94);
        if (i > 0) {
          pdf.addPage([STAGE_WIDTH, STAGE_HEIGHT], 'landscape');
        }
        pdf.addImage(imgData, 'JPEG', 0, 0, STAGE_WIDTH, STAGE_HEIGHT);
      }

      pdf.save('V-Room_Presentation.pdf');
      showToast('✅ V-Room_Presentation.pdf をダウンロードしました', 3200);
    } catch (err) {
      console.error('PDF生成エラー:', err);
      showToast('⚠️ PDF直接生成に失敗したため印刷ダイアログを開きます', 3000);
      window.print();
    } finally {
      stage.style.transform = prevTransform;
      showSlide(originalIndex);
      fitStageToWindow();
      isExporting = false;
    }
  }

  /* =========================================================
     4. PPTXダウンロード生成（X / D キー：編集可能なネイティブPowerPoint生成）
  ========================================================= */
  async function exportToPPTX() {
    if (isExporting) return;
    closePreDeckScreen();

    if (typeof window.PptxGenJS === 'undefined') {
      showToast('⚠️ PptxGenJSライブラリを読み込めませんでした', 3500);
      return;
    }

    isExporting = true;
    showToast('📊 PowerPoint (.pptx) ファイルを生成中...', 0);

    try {
      const pptx = new window.PptxGenJS();
      pptx.layout = 'LAYOUT_16x9';
      pptx.author = '防音VR個室V-Room 企画チーム';
      pptx.company = '第23回 ビジネスプロデュースコンペティション';
      pptx.title = '防音VR個室「V-Room」事業企画プレゼンテーション';

      // 明るいライトテーマ配色
      const C = {
        bg: 'FFFFFF',
        panel: 'F8FAFC',
        border: 'CBD5E1',
        dark: '0F172A',
        sub: '475569',
        text: '1E293B',
        blue: '2563EB',
        cyan: '0284C7',
        green: '059669',
        purple: '7C3AED',
        red: 'DC2626',
        lightBlue: 'EFF6FF'
      };

      const tblOpts = {
        border: { pt: 1, color: C.border },
        fill: C.panel,
        color: C.text,
        fontFace: 'Meiryo',
        fontSize: 9.5,
        valign: 'middle'
      };

      function addHeader(slide, tag, title, pageNum) {
        slide.background = { color: C.bg };
        slide.addText(tag, { x: 0.4, y: 0.2, w: 8.0, h: 0.24, fontSize: 9, bold: true, color: C.blue, fontFace: 'Meiryo' });
        slide.addText(title, { x: 0.4, y: 0.42, w: 8.4, h: 0.4, fontSize: 17, bold: true, color: C.dark, fontFace: 'Meiryo' });
        slide.addText(`${String(pageNum).padStart(2, '0')} / 10`, {
          x: 8.9, y: 0.32, w: 0.7, h: 0.32, fontSize: 10, bold: true, color: C.blue, align: 'center',
          fill: { color: C.lightBlue }, line: { color: C.border, pt: 1 }, fontFace: 'Meiryo'
        });
        slide.addShape(pptx.ShapeType.line, { x: 0.4, y: 0.9, w: 9.2, h: 0, line: { color: C.border, pt: 1.5 } });
      }

      // Slide 1: 表紙
      {
        const s = pptx.addSlide();
        s.background = { color: C.bg };
        s.addText('第23回 ビジネスプロデュースコンペティション（BPC）提案書 ｜ プラン名：防音VR個室V-Room', {
          x: 0.5, y: 0.45, w: 8.5, h: 0.36, fontSize: 11, bold: true, color: C.blue,
          fill: { color: C.lightBlue }, line: { color: C.blue, pt: 1 }, align: 'center', fontFace: 'Meiryo'
        });
        s.addText('日本の狭小な住環境課題と衛生不安を解消する', {
          x: 0.5, y: 1.2, w: 9.0, h: 0.4, fontSize: 16, bold: true, color: C.cyan, fontFace: 'Meiryo'
        });
        s.addText('防音VR個室「V-Room」', {
          x: 0.5, y: 1.6, w: 9.0, h: 0.85, fontSize: 40, bold: true, color: C.dark, fontFace: 'Meiryo'
        });
        s.addText('機材常設・完全防音のVR特化型・有人管理時間貸しレンタルスペース事業', {
          x: 0.5, y: 2.5, w: 9.0, h: 0.4, fontSize: 15, bold: true, color: C.sub, fontFace: 'Meiryo'
        });

        const kpis = [
          { label: '初期投資（1店舗3室）', val: '400万円', sub: '資本金500万（運転資金100万）', col: C.dark },
          { label: '月間営業利益（稼働45%）', val: '44.5万円/月', sub: '月商125.6万 / 利益率35.4%', col: C.cyan },
          { label: '投資回収・損益分岐点', val: '約9.0ヶ月', sub: 'BEP稼働率 28.3% (1日6.8h)', col: C.green },
          { label: '10年後長期ビジョン', val: '年商22.5億円', sub: '全国150拠点（5年後50拠点）', col: C.purple }
        ];
        kpis.forEach((k, idx) => {
          const x = 0.5 + idx * 2.3;
          s.addShape(pptx.ShapeType.roundRect, { x, y: 3.3, w: 2.15, h: 1.7, fill: { color: C.panel }, line: { color: C.blue, pt: 1.5 } });
          s.addText(k.label, { x: x + 0.1, y: 3.42, w: 1.95, h: 0.3, fontSize: 9.5, bold: true, color: C.sub, fontFace: 'Meiryo' });
          s.addText(k.val, { x: x + 0.1, y: 3.8, w: 1.95, h: 0.5, fontSize: 20, bold: true, color: k.col, fontFace: 'Meiryo' });
          s.addText(k.sub, { x: x + 0.1, y: 4.4, w: 1.95, h: 0.4, fontSize: 8.5, color: C.text, fontFace: 'Meiryo' });
        });
      }

      // Slide 2: 背景と課題
      {
        const s = pptx.addSlide();
        addHeader(s, '01 / PROBLEM & INSIGHT（審査項目3）', '背景と課題：世界1位のVR熱量と「日本の住環境・4つの壁」', 2);
        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.05, w: 4.1, h: 2.1, fill: { color: C.lightBlue }, line: { color: C.blue, pt: 1.5 } });
        s.addText('VRChat国別アクセス比率：日本が世界第1位（37%）', { x: 0.55, y: 1.15, w: 3.8, h: 0.4, fontSize: 12.5, bold: true, color: C.blue, fontFace: 'Meiryo' });
        s.addText('・Sensor Tower調査で米国(28%)を抜き世界1位\n・国内メタバース市場は2028年度に1兆8,700億円へ拡大\n・しかし自宅では物理的制約により100%没入できない', {
          x: 0.55, y: 1.6, w: 3.8, h: 1.4, fontSize: 10, color: C.text, fontFace: 'Meiryo'
        });

        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 3.3, w: 4.1, h: 1.95, fill: { color: C.panel }, line: { color: C.border, pt: 1 } });
        s.addText('首都圏単身賃貸 平均20.3㎡ vs 推奨2m×2m空間', { x: 0.55, y: 3.42, w: 3.8, h: 0.35, fontSize: 11.5, bold: true, color: C.dark, fontFace: 'Meiryo' });
        s.addText('総務省「令和5年住宅・土地統計調査」単身平均20.3㎡に対し、SteamVR推奨ルームスケールは家具を除き2m×2m以上が必須。自宅での空間確保は困難。', {
          x: 0.55, y: 3.82, w: 3.8, h: 1.3, fontSize: 9.8, color: C.text, fontFace: 'Meiryo'
        });

        const walls = [
          { t: '壁01：住環境の物理的制約', d: '家具設置後に2m四方を確保できず、壁や家具への衝突・コントローラー破損事故が頻発。' },
          { t: '壁02：深夜の騒音トラブル', d: 'メイン活動時間（21時〜翌3時）のボイスチャットやダンス足音が隣室・家族への騒音に。' },
          { t: '壁03：初期投資30〜40万円の壁', d: 'ハイエンドPC＋HMD＋フルトラセンサー一式の購入費用30〜40万円と高度な設定知識が必要。' },
          { t: '壁04：既存施設の限界と衛生不安', d: 'ネカフェは着座のみ、商業VRは他人の汗・皮脂への嫌悪感、プロスタジオは1.5〜3万円/hと高額。' }
        ];
        walls.forEach((w, idx) => {
          const x = 4.7 + (idx % 2) * 2.5;
          const y = 1.05 + Math.floor(idx / 2) * 2.15;
          s.addShape(pptx.ShapeType.roundRect, { x, y, w: 2.4, h: 2.0, fill: { color: C.panel }, line: { color: C.red, pt: 1.5 } });
          s.addText(w.t, { x: x + 0.12, y: y + 0.14, w: 2.16, h: 0.4, fontSize: 11, bold: true, color: C.red, fontFace: 'Meiryo' });
          s.addText(w.d, { x: x + 0.12, y: y + 0.58, w: 2.16, h: 1.3, fontSize: 9.5, color: C.text, fontFace: 'Meiryo' });
        });
      }

      // Slide 3: ソリューションと料金体系
      {
        const s = pptx.addSlide();
        addHeader(s, '02 / SOLUTION & PRICING（審査項目3・4）', '解決策：「新結合」による防音VR個室と選べる高コスパ料金体系', 3);
        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.02, w: 9.2, h: 0.65, fill: { color: C.lightBlue }, line: { color: C.blue, pt: 1.5 } });
        s.addText('【新結合】 防音クッション個室(Dr-40) × 選べる機材有無(持込/フル貸出) × 買い切りシリコンカバー(500円) ＝ 実質300円/時〜のVR専用空間', {
          x: 0.5, y: 1.1, w: 9.0, h: 0.5, fontSize: 11, bold: true, color: C.blue, align: 'center', fontFace: 'Meiryo'
        });

        const rows = [
          ['部屋タイプ（室数）', 'プラン区分（機材貸出の有無）', '1時間', '3時間パック', '深夜8hパック'],
          ['ライト室（2室）3〜4畳', '① 機器持ち込み（空間＋高速Wi-Fi 6E）', '600円', '1,500円 (500円/h)', '2,400円 (実質300円/h)'],
          ['ライト室（2室）3〜4畳', '② 手ぶらレンタル（Meta Quest 3一式込）', '1,000円', '2,500円 (833円/h)', '3,800円 (実質475円/h)'],
          ['ヘビー室（1室）6〜10畳防音', '③ 機器持ち込み（防音＋PC＋常設BS＋10G）', '1,200円', '3,000円 (1,000円/h)', '4,500円 (実質562円/h)'],
          ['ヘビー室（1室）6〜10畳防音', '④ 手ぶらフルレンタル（PC＋HMD＋フルトラ込）', '2,000円', '5,000円 (1,666円/h)', '6,800円 (実質850円/h)']
        ];
        s.addTable(rows, { ...tblOpts, x: 0.4, y: 1.8, w: 9.2, rowH: 0.52, colW: [2.1, 3.1, 1.0, 1.5, 1.5] });
        s.addText('・オプション：買い切りシリコンカバー 500円/個（原価100円・次回以降マイカバー持参で0円） ／ 単発フルトラッカー追加 +500円/回', {
          x: 0.4, y: 4.75, w: 9.2, h: 0.4, fontSize: 10.5, bold: true, color: C.green, fontFace: 'Meiryo'
        });
      }

      // Slide 4: ビジネススキーム
      {
        const s = pptx.addSlide();
        addHeader(s, '03 / BUSINESS SCHEME（審査項目4）', 'ビジネススキーム：専任スタッフ1名有人管理 × スマートロック省人化', 4);
        const schemeCanvas = document.getElementById('slideCanvasScheme');
        if (schemeCanvas) {
          s.addImage({ data: schemeCanvas.toDataURL('image/png'), x: 0.4, y: 1.05, w: 5.8, h: 4.2 });
        }
        s.addShape(pptx.ShapeType.roundRect, { x: 6.35, y: 1.05, w: 3.25, h: 4.2, fill: { color: C.panel }, line: { color: C.blue, pt: 1.5 } });
        s.addText(
          '① スマートロック自動受付\nWeb予約＋事前決済完了後、利用時間のみ有効な入室＆機材庫PINを自動発行しフロント業務を完全省力化。\n\n' +
          '② 専任スタッフ1名（月26.5万円）\n時給1,260円×154h＋深夜1,575円×22h＝22.9万＋社保・交通費3.6万で有人管理。\n\n' +
          '③ 清掃内製化で外注費ゼロ\nスタッフが日常清掃・除菌・VR保守を直接担い、消耗品実費月1.5万円のみで清潔度100%を維持。',
          { x: 6.5, y: 1.2, w: 2.95, h: 3.9, fontSize: 10, color: C.text, fontFace: 'Meiryo' }
        );
      }

      // Slide 5: 4段階インターバル衛生・機器管理プロトコル
      {
        const s = pptx.addSlide();
        addHeader(s, '04 / HYGIENE & MAINTENANCE PROTOCOL（審査項目4・5）', '利用間の4段階インターバル管理：衛生不安と機器不良を完全払拭', 5);
        const steps = [
          { n: 'STEP 01', t: '予約間の自動15分インターバル', d: '予約システム上で利用枠の間に15分間の入室不可時間を自動挿入し、鉢合わせと熱気残留を遮断。' },
          { n: 'STEP 02', t: '専任スタッフの直接入室検品', d: '①レンズ皮脂・傷確認、②配線ねじれ解消、③トラッカー充電、④境界線ズレ補正、⑤床・壁アルコール除菌。' },
          { n: 'STEP 03', t: 'UV-C＋オゾン殺菌＆急速換気', d: 'スマート保管庫内で波長254nm UV-C照射＋オゾン脱臭(10分)。24h換気＋脱臭機で15分で空気100%入替。' },
          { n: 'STEP 04', t: '肌接触ゼロ＆初心者サポート', d: '個人買い切りシリコンカバー(500円)を必須化し本体への皮脂付着を防止。未経験者には対面装着サポート。' }
        ];
        steps.forEach((st, idx) => {
          const x = 0.4 + idx * 2.33;
          s.addShape(pptx.ShapeType.roundRect, { x, y: 1.05, w: 2.2, h: 2.4, fill: { color: C.panel }, line: { color: C.blue, pt: 1.5 } });
          s.addText(st.n, { x: x + 0.12, y: 1.2, w: 1.96, h: 0.28, fontSize: 9.5, bold: true, color: C.blue, fontFace: 'Meiryo' });
          s.addText(st.t, { x: x + 0.12, y: 1.48, w: 1.96, h: 0.45, fontSize: 11, bold: true, color: C.dark, fontFace: 'Meiryo' });
          s.addText(st.d, { x: x + 0.12, y: 2.0, w: 1.96, h: 1.3, fontSize: 9.2, color: C.text, fontFace: 'Meiryo' });
        });
        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 3.65, w: 9.2, h: 1.6, fill: { color: C.panel }, line: { color: C.border, pt: 1 } });
        s.addText(
          '・技術裏付け：石膏ボード二重貼り＋遮音シート＋50mmウレタン吸音壁で遮音等級Dr-40〜45相当を確保。月間電力1,859kWh（月7.5万円）を精緻に計上。\n' +
          '・24h監視補完：14:00〜23:00コアタイム有人常駐に加え、不在時は退室時AI画像差分チェック（利用中は遮断）＋保管庫重量センサーで24時間監視。',
          { x: 0.55, y: 3.8, w: 8.9, h: 1.3, fontSize: 10, color: C.text, fontFace: 'Meiryo' }
        );
      }

      // Slide 6: 市場規模と競合比較
      {
        const s = pptx.addSlide();
        addHeader(s, '05 / MARKET SIZE & COMPETITIVE ADVANTAGE（審査項目5）', '市場規模（1兆8,700億円予測）と競合ポジショニング比較', 6);
        s.addText('【TAM】1兆8,700億円（2028年度国内メタバース市場予測：矢野経済研究所） ／ 【SAM】525億円 ／ 【SOM】22.5億円（10年後150拠点）', {
          x: 0.4, y: 1.0, w: 9.2, h: 0.4, fontSize: 11, bold: true, color: C.blue, fontFace: 'Meiryo'
        });
        const compRows = [
          ['サービス分類', '空間の可動域・防音', 'VR機材の専門性', '衛生・有人管理', '時間あたり料金帯'],
          ['★当事業「V-Room」', '◎ 3〜10畳 防音クッション', '◎ 持込OK・フルトラ常設', '◎ 有人清掃＋カバー＋UV', '600〜2,000円 (深夜300円/h〜)'],
          ['自宅（単身ワンルーム）', '× 狭小（20.3㎡）衝突危険', '△ 自己負担（30〜40万円）', '◎ 個人専用', '家賃（深夜騒音NG）'],
          ['ネットカフェ（快活CLUB等）', '× 着座のみ（約1畳）', '× 一般PCのみ（VR不可）', '△ 共用機器の簡易清拭', '500〜700円/時'],
          ['商業VR施設', '○ プレイゾーンあり', '△ 指定ゲームのみ（VRC不可）', '△ 不織布マスク（汗付着）', '約3,000円 / 15分'],
          ['プロ用防音撮影スタジオ', '◎ 15畳以上・防音完備', '○ モーションキャプチャ', '○ スタッフ管理', '15,000〜30,000円/時']
        ];
        s.addTable(compRows, { ...tblOpts, x: 0.4, y: 1.5, w: 9.2, rowH: 0.56, colW: [2.0, 1.9, 1.9, 1.8, 1.6] });
      }

      // Slide 7: 初期投資と月次損益
      {
        const s = pptx.addSlide();
        addHeader(s, '06 / FINANCIAL PLAN & UNIT ECONOMICS（審査項目8）', '初期投資計画（400万円）と1店舗あたり基準月次損益（bpc_jigyoushisan2026準拠）', 7);
        const capexRows = [
          ['初期投資項目（Excel列）', '金額（円）'],
          ['資本金（C5）', '5,000,000'],
          ['物件取得費（C7）', '1,500,000'],
          ['設備費（C8：防音110万＋機材100万）', '2,100,000'],
          ['開発費（C9：スマートロック連携）', '200,000'],
          ['開業費（C10：10G回線・採用広報）', '200,000'],
          ['初期投資 計（C15）', '4,000,000'],
          ['初期運転資金（C17）', '1,000,000']
        ];
        s.addTable(capexRows, { ...tblOpts, x: 0.4, y: 1.05, w: 3.6, rowH: 0.48, colW: [2.4, 1.2] });

        const plRows = [
          ['月次損益項目（標準稼働率45%：1日10.8h/室）', '単価 × 数量', '月額（円）'],
          ['売上：ヘビー室（1室）', '1,800円 × 324時間', '583,200'],
          ['売上：ライト室（2室）', '900円 × 648時間', '583,200'],
          ['売上：シリコンカバー・オプション物販', '500円 × 180件', '90,000'],
          ['【月間売上合計（K27）】', '年商換算 15,076,800円', '1,256,400'],
          ['経費：家賃30万＋人件費26.5万（専任1名・社保込）', '固定費（内製清掃）', '▲565,000'],
          ['経費：光熱7.5万＋通信2万＋消耗品1.5万＋積立保険8万', '固定費・保守更新積立', '▲190,000'],
          ['経費：キャッシュレス決済手数料（4.5%）', '変動費', '▲56,538'],
          ['【月間営業利益（V27）】 利益率35.4%', '回収期間 約9.0ヶ月', '+444,862']
        ];
        s.addTable(plRows, { ...tblOpts, x: 4.15, y: 1.05, w: 5.45, rowH: 0.43, colW: [2.65, 1.6, 1.2] });
      }

      // Slide 8: 採算性証明と感度分析
      {
        const s = pptx.addSlide();
        addHeader(s, '07 / PROFITABILITY PROOF & SENSITIVITY（審査項目7・8）', '採算性の定量証明：二階建て収益構造と3シナリオ感度分析', 8);
        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.05, w: 9.2, h: 1.25, fill: { color: C.lightBlue }, line: { color: C.blue, pt: 1.5 } });
        s.addText(
          '① 二階建て採算構造：月間固定費75.5万円に対し、「ヘビー室1室(58.3万円)＋物販(9.0万円)＝67.3万円」だけで固定費の89.2%を回収。\n' +
          '② 低い損益分岐点：1室あたり1日約6.8時間（BEP稼働率28.3%）で黒字化。ライト2室＋ヘビー1室が45㎡内で最短回収（約9.0ヶ月）の最適解。',
          { x: 0.55, y: 1.18, w: 8.9, h: 1.0, fontSize: 10.5, bold: true, color: C.text, fontFace: 'Meiryo' }
        );

        const sensRows = [
          ['シナリオ（稼働率・1日稼働時間）', '月間売上高', '月間総費用（1名人件費込）', '月間営業利益', '投資回収期間'],
          ['悲観シナリオ（稼働率30%：7.2h/日）', '837,600円', '792,692円', '+44,908円（黒字維持）', '約89.1ヶ月'],
          ['★標準シナリオ（稼働率45%：10.8h/日）', '1,256,400円', '811,538円', '+444,862円（利益率35.4%）', '約9.0ヶ月'],
          ['楽観シナリオ（稼働率65%：15.6h/日）', '1,814,800円', '836,666円', '+978,134円（利益率53.9%）', '約4.1ヶ月']
        ];
        s.addTable(sensRows, { ...tblOpts, x: 0.4, y: 2.5, w: 9.2, rowH: 0.62, colW: [2.6, 1.6, 2.0, 1.8, 1.2] });
      }

      // Slide 9: 将来の計画・展望（ロードマップ）
      {
        const s = pptx.addSlide();
        addHeader(s, '08 / GROWTH ROADMAP（審査項目6）', '将来の計画・展望：5年後50拠点・10年後150拠点（年商22.5億円）への拡大', 9);
        const rmRows = [
          ['フェーズ・期間', '展開拠点数', 'グループ年商目標', '主要な実現施策・市場拡大の根拠'],
          ['Phase 1〜2（1〜18ヶ月）', '直営 1店 → 3店', '1,508万 → 4,520万円', '秋葉原1号店を9ヶ月回収後、新宿・池袋へ展開。正社員1名が3店統括。'],
          ['Phase 3（3年目）', '全国 15拠点', '2.25億円', '空きビルオーナー向け「初期投資400万円・1名省人化有人FCパッケージ」展開。'],
          ['Phase 4（5年後目標）', '全国 50拠点', '7.5億円', '国内メタバース市場1兆8,700億円到達を追い風に政令指定都市駅前を網羅。'],
          ['Phase 5（10年後目標）', '全国 150拠点', '22.5億円', '企業のXR遠隔技術研修・自治体のメタバース不登校支援教室（BtoB/BtoG）へ拡張。']
        ];
        s.addTable(rmRows, { ...tblOpts, x: 0.4, y: 1.1, w: 9.2, rowH: 0.78, colW: [2.0, 1.4, 1.6, 4.2] });
      }

      // Slide 10: 法規制適合・リスク管理・総括
      {
        const s = pptx.addSlide();
        addHeader(s, '09 / COMPLIANCE, RISK MANAGEMENT & SUMMARY（審査項目7）', '法規制への完全適合・安全補償体制・ビジネス総括', 10);
        const cards = [
          { t: '風営法・旅館業法対策', d: '・室内からの完全施錠（内鍵固定）を設けない構造\n・防音ドアに防音複層ガラス小窓を配置\n・寝具提供や宿泊を禁止し警察署・保健所へ事前照会' },
          { t: '消防法・W保険・更新積立', d: '・天井不燃材、連動型煙感知器・誘導灯を設置\n・賠償責任保険＋動産総合保険（月1.5万円）加入\n・機器更新積立（月6.5万円）で2〜3年周期の更新を担保' },
          { t: '労働法規遵守と雇用設計', d: '・都最低賃金超（1,260円/h）＋22時以降深夜割増（1,575円/h）\n・週40h上限、労災・雇用・社会保険（約15.5%）完備\n・管理カウンター設置と10分間点検リスト標準化' }
        ];
        cards.forEach((c, idx) => {
          const x = 0.4 + idx * 3.1;
          s.addShape(pptx.ShapeType.roundRect, { x, y: 1.05, w: 3.0, h: 2.35, fill: { color: C.panel }, line: { color: C.blue, pt: 1 } });
          s.addText(c.t, { x: x + 0.14, y: 1.2, w: 2.72, h: 0.38, fontSize: 12, bold: true, color: C.blue, fontFace: 'Meiryo' });
          s.addText(c.d, { x: x + 0.14, y: 1.65, w: 2.72, h: 1.6, fontSize: 9.8, color: C.text, fontFace: 'Meiryo' });
        });

        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 3.6, w: 9.2, h: 1.6, fill: { color: C.lightBlue }, line: { color: C.blue, pt: 2 } });
        s.addText(
          '【総括：防音VR個室「V-Room」】\n' +
          '日本の狭小な住環境と世界1位のVR熱量のギャップを「防音クッション個室×選べる機材有無×有人インターバル清掃×スマートロック省人化」で解決。\n' +
          '初期投資400万円・月間営業利益44.5万円（利益率35.4%）・約9.0ヶ月回収の高い収益性と社会インフラ化を実現します。',
          { x: 0.6, y: 3.75, w: 8.8, h: 1.3, fontSize: 11, bold: true, color: C.dark, fontFace: 'Meiryo' }
        );
      }

      await pptx.writeFile({ fileName: 'V-Room_Presentation.pptx' });
      showToast('✅ V-Room_Presentation.pptx をダウンロードしました', 3200);
    } catch (err) {
      console.error('PPTX生成エラー:', err);
      showToast('⚠️ PPTXファイルの生成中にエラーが発生しました', 3500);
    } finally {
      isExporting = false;
    }
  }

  /* =========================================================
     5. スライド内 Canvas 図表の描画（5種）
  ========================================================= */
  function roundRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  function drawSlideCharts() {
    // 1. VRChat国別アクセス比率 (Slide 2)
    const cShare = document.getElementById('slideChartShare');
    if (cShare) {
      const ctx = cShare.getContext('2d');
      ctx.clearRect(0, 0, cShare.width, cShare.height);
      const items = [
        { label: '日本 (世界1位)', val: 37, col: '#0284c7' },
        { label: '米国 (2位)', val: 28, col: '#2563eb' },
        { label: '英国 (3位)', val: 6, col: '#6366f1' },
        { label: 'その他', val: 29, col: '#64748b' }
      ];
      items.forEach((it, i) => {
        const y = 18 + i * 46;
        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 14px sans-serif';
        ctx.fillText(it.label, 10, y + 18);

        ctx.fillStyle = '#e2e8f0';
        roundRect(ctx, 130, y, 220, 26, 6);
        ctx.fill();

        ctx.fillStyle = it.col;
        roundRect(ctx, 130, y, (220 * it.val) / 45, 26, 6);
        ctx.fill();

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 15px sans-serif';
        ctx.fillText(`${it.val}%`, 360, y + 19);
      });
    }

    // 2. ビジネススキーム図 (Slide 4)
    const cScheme = document.getElementById('slideCanvasScheme');
    if (cScheme) {
      const ctx = cScheme.getContext('2d');
      ctx.clearRect(0, 0, cScheme.width, cScheme.height);

      const topBoxes = [
        { x: 16, w: 230, t: '🏢 物件・10G通信キャリア', s: '駅近14坪・10Gbps光回線', c: '家賃30万・光熱通信9.5万/月' },
        { x: 264, w: 230, t: '🥽 機材・カバー納入元', s: 'PC・Quest3・BS・カバー', c: '初期100万・積立保険8万/月' },
        { x: 512, w: 232, t: '👤 専任スタッフ1名・決済', s: '利用間直接清掃・検品・決済', c: '人件費26.5万・消耗1.5万・4.5%' }
      ];
      topBoxes.forEach((b) => {
        ctx.fillStyle = '#f8fafc';
        ctx.strokeStyle = '#94a3b8';
        ctx.lineWidth = 2;
        roundRect(ctx, b.x, 12, b.w, 105, 10);
        ctx.fill();
        ctx.stroke();

        ctx.fillStyle = '#1d4ed8';
        ctx.font = 'bold 14px sans-serif';
        ctx.fillText(b.t, b.x + 12, 40);
        ctx.fillStyle = '#334155';
        ctx.font = '13px sans-serif';
        ctx.fillText(b.s, b.x + 12, 66);
        ctx.fillStyle = '#b45309';
        ctx.font = 'bold 12px sans-serif';
        ctx.fillText(`◀ ${b.c}`, b.x + 12, 94);
      });

      ctx.fillStyle = '#475569';
      ctx.font = 'bold 13px sans-serif';
      ctx.fillText('▲ 月間総費用 81.2万円（うち固定費75.5万円） ｜ ▼ 空間・10G回線・有人検品済みの清潔な機材供給', 65, 148);

      // 中央コアボックス
      ctx.fillStyle = '#eff6ff';
      ctx.strokeStyle = '#2563eb';
      ctx.lineWidth = 3;
      roundRect(ctx, 16, 168, 728, 115, 14);
      ctx.fill();
      ctx.stroke();

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 19px sans-serif';
      ctx.fillText('🚀 当事業：防音VR個室「V-Room」（1店舗3室：ヘビー1室＋ライト2室）', 45, 205);
      ctx.fillStyle = '#1e3a8a';
      ctx.font = '14px sans-serif';
      ctx.fillText('・Web事前決済＋スマートロックPIN自動発行でフロント受付を省力化（専任スタッフ1名体制）', 45, 236);
      ctx.fillText('・予約間15分インターバルでの直接入室検品・除菌清拭＋UV-C殺菌庫＋買い切りカバー(500円)', 45, 262);

      ctx.fillStyle = '#047857';
      ctx.font = 'bold 13px sans-serif';
      ctx.fillText('▼ 安全・清潔なVR空間と選べる機材の提供 ｜ ▲ 利用料・物販支払（月商125.6万／営業利益44.5万）', 55, 314);

      const botBoxes = [
        { x: 16, w: 354, t: '🧑‍🎓 ライト層・持込層（昼〜夕方中心）', s: '持込600円/h・貸出1,000円/h（深夜実質300円/h〜）', c: 'ライト2室売上：月583,200円（平均900円/h）', col: '#0284c7' },
        { x: 390, w: 354, t: '💃 ヘビー層・配信者層（夜間・深夜・休日）', s: '持込1,200円/h・フル貸出2,000円/h（常設BS完備）', c: 'ヘビー1室＋物販：月673,200円（固定費89.2%回収）', col: '#7c3aed' }
      ];
      botBoxes.forEach((b) => {
        ctx.fillStyle = '#f8fafc';
        ctx.strokeStyle = b.col;
        ctx.lineWidth = 2;
        roundRect(ctx, b.x, 334, b.w, 120, 10);
        ctx.fill();
        ctx.stroke();

        ctx.fillStyle = b.col;
        ctx.font = 'bold 15px sans-serif';
        ctx.fillText(b.t, b.x + 14, 366);
        ctx.fillStyle = '#334155';
        ctx.font = '13px sans-serif';
        ctx.fillText(b.s, b.x + 14, 396);
        ctx.fillStyle = '#047857';
        ctx.font = 'bold 13px sans-serif';
        ctx.fillText(b.c, b.x + 14, 428);
      });
    }

    // 3. 国内メタバース市場規模推移 (Slide 6)
    const cMarket = document.getElementById('slideChartMarket');
    if (cMarket) {
      const ctx = cMarket.getContext('2d');
      ctx.clearRect(0, 0, cMarket.width, cMarket.height);
      const data = [
        { y: '2023', v: 1863 },
        { y: '2024見込', v: 2750 },
        { y: '2025予', v: 4400 },
        { y: '2026予', v: 7500 },
        { y: '2027予', v: 12000 },
        { y: '2028予', v: 18700 }
      ];
      const baseLine = 168;
      const maxH = 125;
      data.forEach((d, i) => {
        const x = 18 + i * 66;
        const bh = Math.max(10, (d.v / 18700) * maxH);
        const y = baseLine - bh;

        ctx.fillStyle = i === 5 ? '#0284c7' : '#2563eb';
        roundRect(ctx, x, y, 42, bh, 5);
        ctx.fill();

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 11px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(d.v >= 10000 ? `${(d.v / 10000).toFixed(2)}兆` : `${d.v}億`, x + 21, y - 6);

        ctx.fillStyle = '#475569';
        ctx.font = 'bold 11px sans-serif';
        ctx.fillText(d.y, x + 21, baseLine + 18);
      });
      ctx.textAlign = 'left';
    }

    // 4. 3シナリオ感度分析 (Slide 8)
    const cSens = document.getElementById('slideChartSensitivity');
    if (cSens) {
      const ctx = cSens.getContext('2d');
      ctx.clearRect(0, 0, cSens.width, cSens.height);
      const groups = [
        { name: '悲観 (稼働30%)', rev: 83.8, exp: 79.3, prof: 4.5 },
        { name: '★標準 (稼働45%)', rev: 125.6, exp: 81.2, prof: 44.5 },
        { name: '楽観 (稼働65%)', rev: 181.5, exp: 83.7, prof: 97.8 }
      ];
      const baseLine = 185;
      const scale = 140 / 190;

      groups.forEach((g, i) => {
        const gx = 35 + i * 165;
        const bars = [
          { val: g.rev, col: '#2563eb', lbl: '売上' },
          { val: g.exp, col: '#dc2626', lbl: '費用' },
          { val: g.prof, col: '#059669', lbl: '利益' }
        ];
        bars.forEach((b, bi) => {
          const bx = gx + bi * 38;
          const bh = Math.max(6, b.val * scale);
          const by = baseLine - bh;
          ctx.fillStyle = b.col;
          roundRect(ctx, bx, by, 30, bh, 4);
          ctx.fill();

          ctx.fillStyle = '#0f172a';
          ctx.font = 'bold 11px sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText(`${b.val}`, bx + 15, by - 5);
        });
        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 12px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(g.name, gx + 53, baseLine + 20);
      });
      ctx.textAlign = 'left';
    }

    // 5. 5年後・10年後ロードマップ (Slide 9)
    const cRoad = document.getElementById('slideChartRoadmap');
    if (cRoad) {
      const ctx = cRoad.getContext('2d');
      ctx.clearRect(0, 0, cRoad.width, cRoad.height);
      const phases = [
        { p: 'Phase1(6m)', stores: 1, rev: '0.15億', h: 18 },
        { p: 'Phase2(18m)', stores: 3, rev: '0.45億', h: 38 },
        { p: 'Phase3(3年)', stores: 15, rev: '2.25億', h: 85 },
        { p: 'Phase4(5年)', stores: 50, rev: '7.5億円', h: 155 },
        { p: 'Phase5(10年)', stores: 150, rev: '22.5億円', h: 235 }
      ];
      const baseLine = 270;
      const pts = [];

      phases.forEach((ph, i) => {
        const x = 28 + i * 98;
        const y = baseLine - ph.h;
        ctx.fillStyle = i >= 3 ? '#0284c7' : '#2563eb';
        roundRect(ctx, x, y, 56, ph.h, 6);
        ctx.fill();

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 12px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(`${ph.stores}拠点`, x + 28, y - 22);
        ctx.fillStyle = '#059669';
        ctx.fillText(ph.rev, x + 28, y - 6);

        ctx.fillStyle = '#334155';
        ctx.font = 'bold 11px sans-serif';
        ctx.fillText(ph.p, x + 28, baseLine + 20);
        pts.push({ x: x + 28, y: y + 10 });
      });

      ctx.strokeStyle = '#d97706';
      ctx.lineWidth = 3;
      ctx.beginPath();
      pts.forEach((pt, idx) => {
        if (idx === 0) ctx.moveTo(pt.x, pt.y);
        else ctx.lineTo(pt.x, pt.y);
      });
      ctx.stroke();
      ctx.textAlign = 'left';
    }
  }

  /* =========================================================
     6. キーボードショートカット制御（日本語IMEオン・物理キーコード完全対応）
  ========================================================= */
  function ensureFocus() {
    window.focus();
    if (document.body) {
      document.body.focus({ preventScroll: true });
    }
  }

  function handleKeyDown(e) {
    if (e.ctrlKey || e.altKey || e.metaKey) return;

    const code = e.code || '';
    const rawKey = e.key || '';
    // 日本語IME全角文字（例：「ｆ」「１」「　」）も半角小文字へ正規化
    const normKey = rawKey.normalize ? rawKey.normalize('NFKC').toLowerCase() : rawKey.toLowerCase();

    const isNext =
      code === 'ArrowRight' || code === 'ArrowDown' || code === 'Space' || code === 'Enter' || code === 'NumpadEnter' || code === 'PageDown' ||
      rawKey === 'ArrowRight' || rawKey === 'ArrowDown' || normKey === ' ' || rawKey === 'Enter' || rawKey === 'PageDown';

    const isPrev =
      code === 'ArrowLeft' || code === 'ArrowUp' || code === 'Backspace' || code === 'PageUp' ||
      rawKey === 'ArrowLeft' || rawKey === 'ArrowUp' || rawKey === 'Backspace' || rawKey === 'PageUp';

    const isFullscreen = code === 'KeyF' || normKey === 'f';
    const isPdf = code === 'KeyP' || normKey === 'p';
    const isPptx = code === 'KeyX' || code === 'KeyD' || normKey === 'x' || normKey === 'd';
    const isIndex = code === 'KeyI' || normKey === 'i';
    const isHome = code === 'Home' || rawKey === 'Home';
    const isEnd = code === 'End' || rawKey === 'End';

    let digitNum = null;
    if (/^Digit[1-9]$/.test(code)) {
      digitNum = parseInt(code.replace('Digit', ''), 10);
    } else if (/^Numpad[1-9]$/.test(code)) {
      digitNum = parseInt(code.replace('Numpad', ''), 10);
    } else if (code === 'Digit0' || code === 'Numpad0' || normKey === '0') {
      digitNum = 10;
    } else if (/^[1-9]$/.test(normKey)) {
      digitNum = parseInt(normKey, 10);
    }

    // プレゼン画面遷移前のスタンバイ画面が表示されている場合
    if (isPreDeckVisible()) {
      if (isFullscreen) {
        e.preventDefault();
        toggleFullscreen();
        showSlide(0);
        return;
      }
      if (isNext) {
        e.preventDefault();
        closePreDeckScreen();
        showSlide(0);
        return;
      }
      if (digitNum !== null) {
        e.preventDefault();
        closePreDeckScreen();
        showSlide(digitNum - 1);
        return;
      }
      if (isPdf) {
        e.preventDefault();
        closePreDeckScreen();
        if (e.shiftKey) window.print();
        else exportToPDF();
        return;
      }
      if (isPptx) {
        e.preventDefault();
        closePreDeckScreen();
        exportToPPTX();
        return;
      }
      if (isIndex) {
        e.preventDefault();
        window.location.href = 'index.html';
        return;
      }
    }

    if (isNext) {
      e.preventDefault();
      nextSlide();
    } else if (isPrev) {
      e.preventDefault();
      prevSlide();
    } else if (isHome) {
      e.preventDefault();
      showSlide(0);
    } else if (isEnd) {
      e.preventDefault();
      showSlide(slides.length - 1);
    } else if (digitNum !== null) {
      e.preventDefault();
      showSlide(digitNum - 1);
    } else if (isFullscreen) {
      e.preventDefault();
      toggleFullscreen();
    } else if (isPdf) {
      e.preventDefault();
      if (e.shiftKey) {
        window.print();
      } else {
        exportToPDF();
      }
    } else if (isPptx) {
      e.preventDefault();
      exportToPPTX();
    } else if (isIndex) {
      e.preventDefault();
      window.location.href = 'index.html';
    }
  }

  /* =========================================================
     7. 初期化（即時セットアップ＋キャプチャフェーズでのキー捕捉）
  ========================================================= */
  function initDeck() {
    fitStageToWindow();
    drawSlideCharts();
    ensureFocus();

    const hashMatch = window.location.hash.match(/^#slide-(\d+)$/);
    if (hashMatch) {
      const idx = parseInt(hashMatch[1], 10) - 1;
      if (idx >= 0 && idx < slides.length) {
        showSlide(idx);
      }
    }
  }

  window.addEventListener('resize', fitStageToWindow);
  window.addEventListener('keydown', handleKeyDown, true);
  window.addEventListener('click', () => {
    ensureFocus();
    if (isPreDeckVisible()) {
      closePreDeckScreen();
      showSlide(0);
    }
  });

  if (document.readyState === 'loading') {
    window.addEventListener('DOMContentLoaded', initDeck);
  } else {
    initDeck();
  }
})();


  function isPreDeckVisible() {
    return preDeckScreen && !preDeckScreen.classList.contains('hidden');
  }


  function closePreDeckScreen() {
    if (preDeckScreen) {
      preDeckScreen.classList.add('hidden');
    }
  }
