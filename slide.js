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
  const preDeckScreen = document.getElementById('pre-deck-screen');

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
        fontSize: 11,
        valign: 'middle'
      };

      const totalPages = slides.length || 14;

      function addHeader(slide, tag, title, pageNum) {
        slide.background = { color: C.bg };
        slide.addText(tag, { x: 0.4, y: 0.2, w: 8.0, h: 0.24, fontSize: 9.5, bold: true, color: C.blue, fontFace: 'Meiryo' });
        slide.addText(title, { x: 0.4, y: 0.42, w: 8.3, h: 0.4, fontSize: 17.5, bold: true, color: C.dark, fontFace: 'Meiryo' });
        slide.addText(`${String(pageNum).padStart(2, '0')} / ${totalPages}`, {
          x: 8.8, y: 0.32, w: 0.8, h: 0.32, fontSize: 10, bold: true, color: C.blue, align: 'center',
          fill: { color: C.lightBlue }, line: { color: C.border, pt: 1 }, fontFace: 'Meiryo'
        });
        slide.addShape(pptx.ShapeType.line, { x: 0.4, y: 0.9, w: 9.2, h: 0, line: { color: C.border, pt: 1.5 } });
      }

      function addFullChartSlide(tag, title, pageNum, canvasId) {
        const s = pptx.addSlide();
        addHeader(s, tag, title, pageNum);
        const c = document.getElementById(canvasId);
        if (c) {
          s.addImage({ data: c.toDataURL('image/png'), x: 0.4, y: 1.02, w: 9.2, h: 4.3 });
        }
      }

      // Slide 1: 表紙
      {
        const s = pptx.addSlide();
        s.background = { color: C.bg };
        s.addText('第23回 ビジネスプロデュースコンペティション（BPC） ｜ プラン名：防音VR個室V-Room', {
          x: 0.5, y: 0.45, w: 8.5, h: 0.36, fontSize: 11, bold: true, color: C.blue,
          fill: { color: C.lightBlue }, line: { color: C.blue, pt: 1 }, align: 'center', fontFace: 'Meiryo'
        });
        s.addText('世界1位のVR熱狂を解放する、日本初の防音・超清潔プライベート空間', {
          x: 0.5, y: 1.2, w: 9.0, h: 0.4, fontSize: 16, bold: true, color: C.cyan, fontFace: 'Meiryo'
        });
        s.addText('防音VR個室「V-Room」', {
          x: 0.5, y: 1.6, w: 9.0, h: 0.85, fontSize: 40, bold: true, color: C.dark, fontFace: 'Meiryo'
        });
        s.addText('入室1分で全身フルトラ起動！ ネカフェ感覚（実質300円/時〜）で遊べる次世代空間', {
          x: 0.5, y: 2.5, w: 9.0, h: 0.4, fontSize: 15, bold: true, color: C.sub, fontFace: 'Meiryo'
        });

        const kpis = [
          { label: '圧倒的コスパ', val: '実質300円/時〜', sub: '手ぶらも持込も自由自在', col: C.cyan },
          { label: '初期投資（14坪・3室）', val: '400万円', sub: '営業利益率35.4%（月44.5万）', col: C.dark },
          { label: '投資回収スピード', val: '約9.0ヶ月', sub: '損益分岐稼働率 わずか28.3%', col: C.green },
          { label: '10年後ビジョン', val: '年商22.5億円', sub: '全国150拠点（5年後50拠点）', col: C.purple }
        ];
        kpis.forEach((k, idx) => {
          const x = 0.5 + idx * 2.3;
          s.addShape(pptx.ShapeType.roundRect, { x, y: 3.3, w: 2.15, h: 1.7, fill: { color: C.panel }, line: { color: C.blue, pt: 1.5 } });
          s.addText(k.label, { x: x + 0.1, y: 3.42, w: 1.95, h: 0.3, fontSize: 9.5, bold: true, color: C.sub, fontFace: 'Meiryo' });
          s.addText(k.val, { x: x + 0.1, y: 3.8, w: 1.95, h: 0.5, fontSize: 20, bold: true, color: k.col, fontFace: 'Meiryo' });
          s.addText(k.sub, { x: x + 0.1, y: 4.4, w: 1.95, h: 0.4, fontSize: 8.5, color: C.text, fontFace: 'Meiryo' });
        });
      }

      // Slide 2: 課題（4つの壁）
      {
        const s = pptx.addSlide();
        addHeader(s, '01 / PROBLEM', 'なぜ自宅でVRを100%楽しめないのか？「4つの壁」', 2);
        const walls = [
          { t: '壁 01 ｜ 空間の壁：部屋が狭すぎて家具に激突！', d: '単身部屋（平均20.3㎡）では動ける2m四方がなく、壁への激突やコントローラー破損が多発。' },
          { t: '壁 02 ｜ 防音の壁：深夜に声も足音も出せない！', d: 'イベント全盛の21時〜翌3時は、ボイスチャットやダンスの足音が隣室・家族への騒音に。' },
          { t: '壁 03 ｜ 費用の壁：機材一式40万円＆設定が難解！', d: '高性能PC＋VRゴーグル＋全身センサーで約40万円。配線や初期設定も初心者には困難。' },
          { t: '壁 04 ｜ 衛生の壁：他人の汗が不快・スタジオは高額！', d: 'ネカフェは着座のみ、商業VRは皮脂汚れが不快、撮影スタジオは1時間2万円〜と高すぎる。' }
        ];
        walls.forEach((w, idx) => {
          const x = 0.4 + (idx % 2) * 4.7;
          const y = 1.1 + Math.floor(idx / 2) * 2.1;
          s.addShape(pptx.ShapeType.roundRect, { x, y, w: 4.5, h: 1.95, fill: { color: C.panel }, line: { color: C.red, pt: 2 } });
          s.addText(w.t, { x: x + 0.18, y: y + 0.18, w: 4.14, h: 0.45, fontSize: 13.5, bold: true, color: C.red, fontFace: 'Meiryo' });
          s.addText(w.d, { x: x + 0.18, y: y + 0.72, w: 4.14, h: 1.05, fontSize: 11.5, color: C.text, fontFace: 'Meiryo' });
        });
      }

      // Slide 3: 【グラフ専用】世界1位のVR熱狂 × 住環境ギャップ
      addFullChartSlide('02 / DATA INSIGHT（グラフ）', '世界第1位のVR熱狂（37%）× 日本の狭小住宅ギャップ', 3, 'slideChartShare');

      // Slide 4: ソリューションと料金プラン
      {
        const s = pptx.addSlide();
        addHeader(s, '03 / SOLUTION & PRICING', 'V-Roomの3大魅力：実質300円/時〜で「完全防音・超清潔」', 4);
        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.02, w: 9.2, h: 0.7, fill: { color: C.lightBlue }, line: { color: C.blue, pt: 1.5 } });
        s.addText('🛡️ 完全防音＆クッション壁 × ⚡ 入室1分でフルトラ起動 × ✨ 肌接触ゼロの超清潔 ＝ 実質300円/時〜！', {
          x: 0.5, y: 1.12, w: 9.0, h: 0.5, fontSize: 12, bold: true, color: C.blue, align: 'center', fontFace: 'Meiryo'
        });

        const rows = [
          ['部屋タイプ', 'プラン（手ぶら／持込）', '1時間', '3時間パック', '深夜8hパック（最安）'],
          ['ライト室（2室）3〜4畳', '① 機器持込（空間＋高速Wi-Fi）', '600円', '1,500円', '2,400円（実質300円/h）'],
          ['ライト室（2室）3〜4畳', '② 手ぶら体験（Quest 3一式込）', '1,000円', '2,500円', '3,800円（実質475円/h）'],
          ['ヘビー室（1室）6〜10畳防音', '③ 機器持込（防音＋PC＋常設BS）', '1,200円', '3,000円', '4,500円（実質562円/h）'],
          ['ヘビー室（1室）6〜10畳防音', '④ 手ぶらフルトラ（全機材込）', '2,000円', '5,000円', '6,800円（実質850円/h）']
        ];
        s.addTable(rows, { ...tblOpts, x: 0.4, y: 1.9, w: 9.2, rowH: 0.62, colW: [2.2, 2.8, 1.1, 1.3, 1.8] });
      }

      // Slide 5: 【図解専用】ビジネススキーム図
      addFullChartSlide('04 / BUSINESS SCHEME（図解）', 'ビジネススキーム：スマートロック無人受付 × 専任1名管理', 5, 'slideCanvasScheme');

      // Slide 6: 4段階の超・清潔リセット＆運営の強み
      {
        const s = pptx.addSlide();
        addHeader(s, '05 / HYGIENE & OPERATION', '女性も初心者も安心！ 利用間15分の「4段階・超清潔リセット」', 6);
        const steps = [
          { n: 'STEP 01', t: '自動15分インターバル', d: '予約間に15分の空白を自動確保。鉢合わせと熱気を完全遮断。' },
          { n: 'STEP 02', t: 'プロが毎回入室リセット', d: 'スタッフが毎回入室し、レンズ・配線・充電・床除菌を徹底。' },
          { n: 'STEP 03', t: 'UV-C殺菌＆急速換気', d: '紫外線＋オゾン庫で99.9%除菌。15分で空気を100%入替。' },
          { n: 'STEP 04', t: 'マイカバーで接触ゼロ', d: '個人カバー(500円)で汗付着ゼロ。次回持参でずっと0円！' }
        ];
        steps.forEach((st, idx) => {
          const x = 0.4 + idx * 2.33;
          s.addShape(pptx.ShapeType.roundRect, { x, y: 1.05, w: 2.2, h: 2.3, fill: { color: C.panel }, line: { color: C.blue, pt: 1.5 } });
          s.addText(st.n, { x: x + 0.12, y: 1.18, w: 1.96, h: 0.3, fontSize: 10, bold: true, color: C.blue, fontFace: 'Meiryo' });
          s.addText(st.t, { x: x + 0.12, y: 1.5, w: 1.96, h: 0.45, fontSize: 12, bold: true, color: C.dark, fontFace: 'Meiryo' });
          s.addText(st.d, { x: x + 0.12, y: 2.0, w: 1.96, h: 1.2, fontSize: 10.5, color: C.text, fontFace: 'Meiryo' });
        });
        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 3.55, w: 9.2, h: 1.7, fill: { color: C.lightBlue }, line: { color: C.blue, pt: 1.5 } });
        s.addText(
          '🔑 無人受付 × 有人ケア：スマートロックでレジ人件費ゼロ。専任1名は「清潔＆初心者サポート」に100%集中。\n' +
          '🛡️ Dr-40防音 × 10Gbps回線 × 24h保護：深夜の熱唱も安心。不在時も退室時AI差分（利用中は完全遮断）で安全保護。',
          { x: 0.55, y: 3.7, w: 8.9, h: 1.4, fontSize: 11.5, bold: true, color: C.text, fontFace: 'Meiryo' }
        );
      }

      // Slide 7: 【グラフ専用】国内メタバース市場予測＆TAM/SAM/SOM
      addFullChartSlide('06 / MARKET GROWTH（グラフ）', '国内市場は4年で約6.8倍（1兆8,700億円）へ急成長', 7, 'slideChartMarket');

      // Slide 8: 競合比較マトリクス
      {
        const s = pptx.addSlide();
        addHeader(s, '07 / COMPETITIVE ADVANTAGE', '競合比較：すべての不満を解消する「4冠独占」ポジション', 8);
        const compRows = [
          ['サービス分類', '広さ・防音性', 'VR・フルトラ環境', '清潔さ・安心感', '料金の手頃さ'],
          ['★防音VR個室「V-Room」', '◎ 防音クッション壁', '◎ 入室1分フルトラ', '◎ 有人除菌＋カバー', '◎ 実質300円/h〜'],
          ['自宅（ワンルーム）', '× 狭小・家具に激突', '△ 初期費用40万円', '◎ 自分専用', '△ 深夜の騒音NG'],
          ['ネットカフェ個室', '× 着座のみ（1畳）', '× VR機器なし', '△ 簡易清掃のみ', '◎ 500円/h〜'],
          ['商業VRアトラクション', '○ プレイ空間あり', '× VRChat利用不可', '× 他人の汗が不快', '× 約3,000円/15分'],
          ['プロ用撮影スタジオ', '◎ 大空間・防音', '○ 専門機材完備', '○ スタッフ管理', '× 1.5万〜3万円/h']
        ];
        s.addTable(compRows, { ...tblOpts, fontSize: 11.5, x: 0.4, y: 1.05, w: 9.2, rowH: 0.62, colW: [2.2, 1.8, 1.8, 1.8, 1.6] });
        s.addText('💡 「完全防音」「フルトラ常設」「有人超清潔」「実質300円/時〜」の4条件を同時に満たす競合はゼロ！', {
          x: 0.4, y: 4.85, w: 9.2, h: 0.4, fontSize: 12, bold: true, color: C.blue, fontFace: 'Meiryo'
        });
      }

      // Slide 9: 初期投資400万円と月次収支モデル
      {
        const s = pptx.addSlide();
        addHeader(s, '08 / UNIT ECONOMICS', '投資魅力：初期投資400万円・利益率35.4%・約9ヶ月でスピード回収', 9);
        const capexRows = [
          ['初期投資項目', '金額'],
          ['物件取得費（駅近14坪）', '150万円'],
          ['防音施工・VR機材', '210万円'],
          ['スマートロック・開業費', '40万円'],
          ['★初期投資 合計（3室）', '400万円'],
          ['運転資金（資本金500万）', '100万円']
        ];
        s.addTable(capexRows, { ...tblOpts, x: 0.4, y: 1.05, w: 3.6, rowH: 0.62, colW: [2.4, 1.2] });

        const plRows = [
          ['1店舗 月次収支（稼働45%）', '内訳', '月額'],
          ['ヘビー室（1室）', '1,800円/h × 324h', '58.3万円'],
          ['ライト室（2室）', '900円/h × 648h', '58.3万円'],
          ['物販・オプション', 'カバー・追加機器', '9.0万円'],
          ['【月間売上高 合計】', '年商 1,508万円', '125.6万円'],
          ['家賃・1名人件費', '家賃30万＋人件費26.5万', '▲56.5万円'],
          ['光熱・通信・積立他', '光熱通信9.5万＋積立等15.2万', '▲24.7万円'],
          ['【月間営業利益】', '利益率35.4%（9ヶ月回収）', '+44.5万円']
        ];
        s.addTable(plRows, { ...tblOpts, fontSize: 10.5, x: 4.15, y: 1.05, w: 5.45, rowH: 0.5, colW: [2.2, 2.0, 1.25] });
      }

      // Slide 10: なぜ3室で儲かるのか？（二階建て収益構造）
      {
        const s = pptx.addSlide();
        addHeader(s, '09 / PROFIT STRUCTURE', 'なぜ3室で儲かるのか？ ヘビー1室で固定費9割を賄う「二階建て構造」', 10);
        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 1.05, w: 9.2, h: 1.35, fill: { color: C.lightBlue }, line: { color: C.blue, pt: 1.5 } });
        s.addText(
          '🏢 1階：月間固定費75.5万円に対し、「ヘビー1室＋物販（月67.3万円）」だけで固定費の89.2%を回収！\n' +
          '🚀 2階：固定費のほぼ全てをカバーするため、「ライト2室の売上（月58.3万円）」がそのまま高利益（+44.5万円）へ！',
          { x: 0.55, y: 1.18, w: 8.9, h: 1.1, fontSize: 11.5, bold: true, color: C.text, fontFace: 'Meiryo' }
        );
        const roomRows = [
          ['14坪内の部屋構成案', '初期投資', '月間営業利益', '回収スピード・特徴'],
          ['案A：ライト室のみ 4室', '360万円', '24.2万円', '約14.9ヶ月（深夜の高単価フルトラ需要を取り逃す）'],
          ['案B：ヘビー室のみ 2室', '480万円', '31.5万円', '約15.2ヶ月（防音・PC投資が重く昼の体験層を逃す）'],
          ['★当事業：ヘビー1室＋ライト2室', '400万円', '+44.5万円', '★約9.0ヶ月（昼夜フル稼働・投資効率ナンバーワン！）']
        ];
        s.addTable(roomRows, { ...tblOpts, x: 0.4, y: 2.6, w: 9.2, rowH: 0.62, colW: [2.5, 1.3, 1.5, 3.9] });
      }

      // Slide 11: 【グラフ専用】3シナリオ感度分析＆損益分岐点
      addFullChartSlide('10 / SENSITIVITY ANALYSIS（グラフ）', '感度分析：損益分岐点はわずか28.3%！ 稼働30%でも黒字維持', 11, 'slideChartSensitivity');

      // Slide 12: 【グラフ専用】5年後・10年後 成長ロードマップグラフ
      addFullChartSlide('11 / GROWTH CHART（グラフ）', '成長予測：全国150拠点・グループ年商22.5億円への拡大軌道', 12, 'slideChartRoadmap');

      // Slide 13: 成長ロードマップ詳細
      {
        const s = pptx.addSlide();
        addHeader(s, '12 / ROADMAP STRATEGY', '展開ステップ：省人化FCパッケージで全国の空きビルを再生', 13);
        const rmRows = [
          ['フェーズ', '展開規模', 'グループ年商', '主な成長アクション'],
          ['Phase 1〜2（〜18ヶ月）', '直営 3店舗', '4,520万円', '秋葉原1号店を9ヶ月で回収し新宿・池袋へ展開。1名で3店を統括。'],
          ['Phase 3（3年目）', '全国 15拠点', '2.25億円', '遊休ビルオーナー向けに「初期400万円・1名運営FC」を提供し東名阪へ拡大。'],
          ['Phase 4（5年後目標）', '全国 50拠点', '7.5億円', '累計会員3万人達成。大型VRイベント連動の公式オフ会拠点としてブランド確立。'],
          ['Phase 5（10年後目標）', '全国 150拠点', '22.5億円', '夜間エンタメに加え、昼間を「企業XR研修」「自治体メタバース不登校支援」へ開放。']
        ];
        s.addTable(rmRows, { ...tblOpts, x: 0.4, y: 1.1, w: 9.2, rowH: 0.8, colW: [2.2, 1.4, 1.5, 4.1] });
      }

      // Slide 14: 法規制クリア・リスク管理・総括
      {
        const s = pptx.addSlide();
        addHeader(s, '13 / RISK MANAGEMENT & SUMMARY', '万全の法規制クリア・無借金更新サイクルと事業総括', 14);
        const cards = [
          { t: '⚖️ 風営法・旅館業法クリア', d: '・内鍵固定を設けず防音ガラス小窓を配置\n・寝具・宿泊を禁止し時間貸し空間として適法運営' },
          { t: '🔄 W保険＆無借金で機材更新', d: '・賠償＋動産保険(月1.5万)で怪我や破損を完全補償\n・毎月6.5万円を経費積立し無借金で最新機材へ刷新' },
          { t: '👥 高待遇＆10分マニュアル', d: '・都最低賃金超(1,260円/h＋深夜割増)と社保完備\n・10分チェックリストで多店舗FCでもプロ品質を再現' }
        ];
        cards.forEach((c, idx) => {
          const x = 0.4 + idx * 3.1;
          s.addShape(pptx.ShapeType.roundRect, { x, y: 1.05, w: 3.0, h: 2.3, fill: { color: C.panel }, line: { color: C.blue, pt: 1.5 } });
          s.addText(c.t, { x: x + 0.14, y: 1.2, w: 2.72, h: 0.4, fontSize: 13, bold: true, color: C.blue, fontFace: 'Meiryo' });
          s.addText(c.d, { x: x + 0.14, y: 1.7, w: 2.72, h: 1.5, fontSize: 11, color: C.text, fontFace: 'Meiryo' });
        });

        s.addShape(pptx.ShapeType.roundRect, { x: 0.4, y: 3.55, w: 9.2, h: 1.65, fill: { color: C.lightBlue }, line: { color: C.blue, pt: 2 } });
        s.addText(
          '【CONCLUSION：世界一のVR熱狂を、日本の新しい「街角インフラ」へ】\n' +
          '「完全防音 × 入室1分フルトラ × 有人超清潔 × 300円/時〜」でユーザーの夢を叶え、\n' +
          '「初期投資400万円・利益率35.4%・9ヶ月回収」の圧倒的なビジネス魅力を実現します。',
          { x: 0.6, y: 3.7, w: 8.8, h: 1.35, fontSize: 12, bold: true, color: C.dark, fontFace: 'Meiryo' }
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
    // 1. VRChat国別シェア ＆ 首都圏単身住宅ギャップ (Slide 3: 1420x620)
    const cShare = document.getElementById('slideChartShare');
    if (cShare) {
      const ctx = cShare.getContext('2d');
      ctx.clearRect(0, 0, cShare.width, cShare.height);

      // 左枠：VRChat国別アクセス比率
      ctx.fillStyle = '#f8fafc';
      ctx.strokeStyle = '#cbd5e1';
      ctx.lineWidth = 2;
      roundRect(ctx, 16, 16, 680, 588, 16);
      ctx.fill();
      ctx.stroke();

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 28px sans-serif';
      ctx.fillText('🌐 VRChat 国別利用シェア', 44, 66);
      ctx.fillStyle = '#0284c7';
      ctx.font = 'bold 20px sans-serif';
      ctx.fillText('日本が米国を抜き「世界第1位」（Sensor Tower調査）', 44, 102);

      const items = [
        { label: '🇯🇵 日本 (1位)', val: 37, col: '#0284c7' },
        { label: '🇺🇸 米国 (2位)', val: 28, col: '#2563eb' },
        { label: '🇬🇧 英国 (3位)', val: 6, col: '#6366f1' },
        { label: '🌍 その他', val: 29, col: '#64748b' }
      ];
      items.forEach((it, i) => {
        const y = 145 + i * 105;
        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 24px sans-serif';
        ctx.fillText(it.label, 44, y + 36);

        ctx.fillStyle = '#e2e8f0';
        roundRect(ctx, 235, y, 310, 52, 10);
        ctx.fill();

        ctx.fillStyle = it.col;
        roundRect(ctx, 235, y, (310 * it.val) / 42, 52, 10);
        ctx.fill();

        ctx.fillStyle = i === 0 ? '#0284c7' : '#0f172a';
        ctx.font = i === 0 ? 'bold 32px sans-serif' : 'bold 26px sans-serif';
        ctx.fillText(`${it.val}%`, 560, y + 38);
      });

      // 右枠：住環境ギャップ
      ctx.fillStyle = '#f8fafc';
      ctx.strokeStyle = '#cbd5e1';
      ctx.lineWidth = 2;
      roundRect(ctx, 724, 16, 680, 588, 16);
      ctx.fill();
      ctx.stroke();

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 28px sans-serif';
      ctx.fillText('🏠 単身部屋の広さ vs VR必要空間', 752, 66);
      ctx.fillStyle = '#dc2626';
      ctx.font = 'bold 20px sans-serif';
      ctx.fillText('平均20.3㎡では家具を置くと動ける空間がない！', 752, 102);

      const rooms = [
        { label: '首都圏単身平均', val: 20.3, note: '× 家具で埋まり衝突', col: '#dc2626' },
        { label: '快適VR必要部屋', val: 30.0, note: '○ 2m×2m確保ライン', col: '#d97706' },
        { label: '国交省誘導水準', val: 40.0, note: '◎ ゆとりある単身室', col: '#059669' }
      ];
      rooms.forEach((rm, i) => {
        const y = 155 + i * 115;
        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 22px sans-serif';
        ctx.fillText(rm.label, 752, y + 30);
        ctx.fillStyle = rm.col;
        ctx.font = 'bold 18px sans-serif';
        ctx.fillText(rm.note, 752, y + 58);

        ctx.fillStyle = '#e2e8f0';
        roundRect(ctx, 965, y + 6, 290, 52, 10);
        ctx.fill();

        ctx.fillStyle = rm.col;
        roundRect(ctx, 965, y + 6, (290 * rm.val) / 45, 52, 10);
        ctx.fill();

        ctx.fillStyle = rm.col;
        ctx.font = 'bold 28px sans-serif';
        ctx.fillText(`${rm.val}㎡`, 1270, y + 42);
      });

      ctx.fillStyle = '#eff6ff';
      ctx.strokeStyle = '#2563eb';
      ctx.lineWidth = 2;
      roundRect(ctx, 752, 505, 624, 72, 12);
      ctx.fill();
      ctx.stroke();
      ctx.fillStyle = '#1e3a8a';
      ctx.font = 'bold 22px sans-serif';
      ctx.fillText('💡 世界1位の熱量 × 狭い自宅 ＝「V-Room」へ需要爆発！', 774, 550);
    }

    // 2. ビジネススキーム図 (Slide 5: 1420x620)
    const cScheme = document.getElementById('slideCanvasScheme');
    if (cScheme) {
      const ctx = cScheme.getContext('2d');
      ctx.clearRect(0, 0, cScheme.width, cScheme.height);

      const topBoxes = [
        { x: 20, w: 435, t: '🏢 物件・10G通信キャリア', s: '駅近14坪ビル・10Gbps光回線', c: '家賃30万・光熱通信9.5万/月' },
        { x: 492, w: 435, t: '🥽 VR機材・カバー納入元', s: 'PC・Quest 3・常設BS・シリコンカバー', c: '初期100万・更新積立＆保険8万/月' },
        { x: 965, w: 435, t: '👤 専任スタッフ1名・決済', s: '利用間15分除菌・検品・事前決済', c: '人件費26.5万・消耗1.5万・手数料4.5%' }
      ];
      topBoxes.forEach((b) => {
        ctx.fillStyle = '#f8fafc';
        ctx.strokeStyle = '#94a3b8';
        ctx.lineWidth = 2.5;
        roundRect(ctx, b.x, 16, b.w, 145, 14);
        ctx.fill();
        ctx.stroke();

        ctx.fillStyle = '#1d4ed8';
        ctx.font = 'bold 24px sans-serif';
        ctx.fillText(b.t, b.x + 20, 56);
        ctx.fillStyle = '#334155';
        ctx.font = 'bold 19px sans-serif';
        ctx.fillText(b.s, b.x + 20, 95);
        ctx.fillStyle = '#b45309';
        ctx.font = 'bold 20px sans-serif';
        ctx.fillText(`◀ ${b.c}`, b.x + 20, 136);
      });

      ctx.fillStyle = '#334155';
      ctx.font = 'bold 21px sans-serif';
      ctx.fillText('▲ 月間総費用 81.2万円（うち固定費75.5万円）　｜　▼ 防音空間・10G回線・有人検品済みの超清潔機材', 160, 198);

      // 中央コアボックス
      ctx.fillStyle = '#eff6ff';
      ctx.strokeStyle = '#2563eb';
      ctx.lineWidth = 4;
      roundRect(ctx, 20, 220, 1380, 165, 18);
      ctx.fill();
      ctx.stroke();

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 30px sans-serif';
      ctx.fillText('🚀 防音VR個室「V-Room」（1店舗3室：完全防音ヘビー1室＋ライト2室）', 55, 272);
      ctx.fillStyle = '#1e3a8a';
      ctx.font = 'bold 22px sans-serif';
      ctx.fillText('✔ スマートロック無人受付でレジ業務ゼロ ➔ スタッフ1名が「清潔＆初心者サポート」に集中！', 55, 318);
      ctx.fillText('✔ 15分有人リセット ＋ UV-C殺菌庫 ＋ マイカバー(500円) ＝「肌接触ゼロの超・清潔空間」', 55, 358);

      ctx.fillStyle = '#047857';
      ctx.font = 'bold 21px sans-serif';
      ctx.fillText('▼ 入室1分でフルトラ起動・選べる機材プラン提供　｜　▲ 利用料・物販支払（月商125.6万 ／ 営業利益+44.5万）', 135, 422);

      const botBoxes = [
        { x: 20, w: 670, t: '🧑‍🎓 ライト層・持込層（昼〜夕方メイン）', s: '持込600円/h・手ぶら1,000円/h（深夜実質300円/h〜）', c: 'ライト2室売上：月58.3万円（平均900円/h）', col: '#0284c7' },
        { x: 730, w: 670, t: '💃 ヘビー層・配信者層（夜間・深夜・休日）', s: '持込1,200円/h・手ぶらフルトラ2,000円/h（常設BS）', c: 'ヘビー1室＋物販：月67.3万円（固定費の89.2%回収！）', col: '#7c3aed' }
      ];
      botBoxes.forEach((b) => {
        ctx.fillStyle = '#f8fafc';
        ctx.strokeStyle = b.col;
        ctx.lineWidth = 3;
        roundRect(ctx, b.x, 445, b.w, 158, 14);
        ctx.fill();
        ctx.stroke();

        ctx.fillStyle = b.col;
        ctx.font = 'bold 25px sans-serif';
        ctx.fillText(b.t, b.x + 24, 490);
        ctx.fillStyle = '#334155';
        ctx.font = 'bold 20px sans-serif';
        ctx.fillText(b.s, b.x + 24, 532);
        ctx.fillStyle = '#047857';
        ctx.font = 'bold 22px sans-serif';
        ctx.fillText(b.c, b.x + 24, 576);
      });
    }

    // 3. 国内メタバース市場規模推移 ＆ TAM/SAM/SOM (Slide 7: 1420x620)
    const cMarket = document.getElementById('slideChartMarket');
    if (cMarket) {
      const ctx = cMarket.getContext('2d');
      ctx.clearRect(0, 0, cMarket.width, cMarket.height);

      // 左枠：市場成長棒グラフ
      ctx.fillStyle = '#f8fafc';
      ctx.strokeStyle = '#cbd5e1';
      ctx.lineWidth = 2;
      roundRect(ctx, 16, 16, 860, 588, 16);
      ctx.fill();
      ctx.stroke();

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 28px sans-serif';
      ctx.fillText('📊 国内メタバース市場規模推移（矢野経済研究所）', 44, 64);
      ctx.fillStyle = '#0284c7';
      ctx.font = 'bold 22px sans-serif';
      ctx.fillText('2024年度 2,750億円 ➔ 2028年度 1兆8,700億円（約6.8倍へ急拡大！）', 44, 102);

      const data = [
        { y: '2023', v: 1863, lbl: '1,863億' },
        { y: '2024見込', v: 2750, lbl: '2,750億' },
        { y: '2025予', v: 4400, lbl: '4,400億' },
        { y: '2026予', v: 7500, lbl: '7,500億' },
        { y: '2027予', v: 12000, lbl: '1.20兆' },
        { y: '2028予', v: 18700, lbl: '1.87兆円' }
      ];
      const baseLine = 535;
      const maxH = 350;
      data.forEach((d, i) => {
        const x = 52 + i * 134;
        const bh = Math.max(24, (d.v / 18700) * maxH);
        const y = baseLine - bh;

        ctx.fillStyle = i === 5 ? '#0284c7' : '#2563eb';
        roundRect(ctx, x, y, 86, bh, 10);
        ctx.fill();

        ctx.fillStyle = i === 5 ? '#0284c7' : '#0f172a';
        ctx.font = i === 5 ? 'bold 24px sans-serif' : 'bold 20px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(d.lbl, x + 43, y - 14);

        ctx.fillStyle = '#334155';
        ctx.font = 'bold 20px sans-serif';
        ctx.fillText(d.y, x + 43, baseLine + 36);
      });
      ctx.textAlign = 'left';

      // 右枠：TAM / SAM / SOM
      const tams = [
        { tag: 'TAM（国内メタバース市場全体）', val: '1兆8,700億円', sub: '2028年度予測（年平均成長率 61.5%）', col: '#2563eb', bg: '#f8fafc' },
        { tag: 'SAM（国内VR空間・機材利用市場）', val: '525億円', sub: '国内関心層350万人 × 年1.5万円支出', col: '#0284c7', bg: '#f8fafc' },
        { tag: 'SOM（V-Room 獲得目標市場）', val: '22.5億円', sub: '10年後 全国150拠点（5年後50拠点 7.5億）', col: '#059669', bg: '#ecfdf5' }
      ];
      tams.forEach((tm, idx) => {
        const y = 16 + idx * 200;
        ctx.fillStyle = tm.bg;
        ctx.strokeStyle = tm.col;
        ctx.lineWidth = 3;
        roundRect(ctx, 900, y, 504, 184, 16);
        ctx.fill();
        ctx.stroke();

        ctx.fillStyle = tm.col;
        ctx.font = 'bold 21px sans-serif';
        ctx.fillText(tm.tag, 928, y + 46);
        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 42px sans-serif';
        ctx.fillText(tm.val, 928, y + 110);
        ctx.fillStyle = '#334155';
        ctx.font = 'bold 19px sans-serif';
        ctx.fillText(tm.sub, 928, y + 154);
      });
    }

    // 4. 3シナリオ感度分析 ＆ 損益分岐点 (Slide 11: 1420x620)
    const cSens = document.getElementById('slideChartSensitivity');
    if (cSens) {
      const ctx = cSens.getContext('2d');
      ctx.clearRect(0, 0, cSens.width, cSens.height);

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 26px sans-serif';
      ctx.fillText('📊 稼働率別 3シナリオ月次収支比較（単位：万円）', 36, 48);

      // 凡例
      const legends = [
        { lbl: '■ 月間売上高', col: '#2563eb', x: 760 },
        { lbl: '■ 月間総費用（1名人件費込）', col: '#dc2626', x: 960 },
        { lbl: '■ 月間営業利益', col: '#059669', x: 1250 }
      ];
      legends.forEach((lg) => {
        ctx.fillStyle = lg.col;
        ctx.font = 'bold 20px sans-serif';
        ctx.fillText(lg.lbl, lg.x, 48);
      });

      const groups = [
        { name: '悲観シナリオ（稼働30%：7.2h/日）', rev: 83.8, exp: 79.3, prof: 4.5, badge: '悪条件でも毎月+4.5万円の黒字維持' },
        { name: '★標準シナリオ（稼働45%：10.8h/日）', rev: 125.6, exp: 81.2, prof: 44.5, badge: '利益率35.4%・約9.0ヶ月でスピード回収' },
        { name: '楽観シナリオ（稼働65%：15.6h/日）', rev: 181.5, exp: 83.7, prof: 97.8, badge: '利益率53.9%・約4.1ヶ月で超早期回収' }
      ];
      const baseLine = 485;
      const scale = 350 / 190;

      // 固定費ライン (75.5万円)
      const bepY = baseLine - 75.5 * scale;
      ctx.strokeStyle = '#d97706';
      ctx.lineWidth = 2.5;
      ctx.setLineDash([10, 6]);
      ctx.beginPath();
      ctx.moveTo(40, bepY);
      ctx.lineTo(1380, bepY);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = '#b45309';
      ctx.font = 'bold 19px sans-serif';
      ctx.fillText('◀ 損益分岐点（BEP）固定費ライン：75.5万円（稼働率28.3%・1日6.8h）', 680, bepY - 12);

      groups.forEach((g, i) => {
        const gx = 75 + i * 455;
        const bars = [
          { val: g.rev, col: '#2563eb', prefix: '' },
          { val: g.exp, col: '#dc2626', prefix: '' },
          { val: g.prof, col: '#059669', prefix: '+' }
        ];
        bars.forEach((b, bi) => {
          const bx = gx + bi * 115;
          const bh = Math.max(16, b.val * scale);
          const by = baseLine - bh;
          ctx.fillStyle = b.col;
          roundRect(ctx, bx, by, 92, bh, 8);
          ctx.fill();

          ctx.fillStyle = b.col;
          ctx.font = 'bold 24px sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText(`${b.prefix}${b.val}万`, bx + 46, by - 14);
        });

        ctx.fillStyle = i === 1 ? '#1d4ed8' : '#0f172a';
        ctx.font = 'bold 22px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(g.name, gx + 160, baseLine + 38);

        ctx.fillStyle = i === 1 ? '#eff6ff' : '#f8fafc';
        ctx.strokeStyle = i === 1 ? '#2563eb' : '#cbd5e1';
        ctx.lineWidth = 2;
        roundRect(ctx, gx - 10, baseLine + 54, 340, 48, 10);
        ctx.fill();
        ctx.stroke();

        ctx.fillStyle = i === 1 ? '#1d4ed8' : '#047857';
        ctx.font = 'bold 18px sans-serif';
        ctx.fillText(g.badge, gx + 160, baseLine + 85);
      });
      ctx.textAlign = 'left';
    }

    // 5. 5年後・10年後ロードマップ (Slide 12: 1420x620)
    const cRoad = document.getElementById('slideChartRoadmap');
    if (cRoad) {
      const ctx = cRoad.getContext('2d');
      ctx.clearRect(0, 0, cRoad.width, cRoad.height);

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 28px sans-serif';
      ctx.fillText('🚀 拠点拡大（棒）＆ グループ年商（折れ線）成長ロードマップ', 44, 52);

      const phases = [
        { p: 'Phase 1 (6ヶ月)', sub: '秋葉原1号店実証', stores: '1拠点', rev: '年商 1,508万', h: 40 },
        { p: 'Phase 2 (18ヶ月)', sub: '首都圏ドミナント', stores: '3拠点', rev: '年商 4,520万', h: 85 },
        { p: 'Phase 3 (3年目)', sub: '1名運営FC展開', stores: '15拠点', rev: '年商 2.25億円', h: 165 },
        { p: 'Phase 4 (5年後)', sub: '全国主要都市網羅', stores: '50拠点', rev: '年商 7.5億円', h: 270 },
        { p: 'Phase 5 (10年後)', sub: 'XR社会インフラ化', stores: '150拠点', rev: '年商 22.5億円', h: 380 }
      ];
      const baseLine = 505;
      const pts = [];

      phases.forEach((ph, i) => {
        const x = 75 + i * 268;
        const y = baseLine - ph.h;
        ctx.fillStyle = i >= 3 ? '#0284c7' : '#2563eb';
        roundRect(ctx, x, y, 136, ph.h, 12);
        ctx.fill();

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 28px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(ph.stores, x + 68, y - 42);

        ctx.fillStyle = '#059669';
        ctx.font = 'bold 24px sans-serif';
        ctx.fillText(ph.rev, x + 68, y - 12);

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 22px sans-serif';
        ctx.fillText(ph.p, x + 68, baseLine + 38);

        ctx.fillStyle = '#475569';
        ctx.font = 'bold 18px sans-serif';
        ctx.fillText(ph.sub, x + 68, baseLine + 68);

        pts.push({ x: x + 68, y: y + 18 });
      });

      ctx.strokeStyle = '#d97706';
      ctx.lineWidth = 5;
      ctx.beginPath();
      pts.forEach((pt, idx) => {
        if (idx === 0) ctx.moveTo(pt.x, pt.y);
        else ctx.lineTo(pt.x, pt.y);
      });
      ctx.stroke();

      pts.forEach((pt) => {
        ctx.fillStyle = '#ffffff';
        ctx.strokeStyle = '#d97706';
        ctx.lineWidth = 4;
        ctx.beginPath();
        ctx.arc(pt.x, pt.y, 9, 0, Math.PI * 2);
        ctx.fill();
        ctx.stroke();
      });
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

  function isPreDeckVisible() {
    return Boolean(preDeckScreen && !preDeckScreen.classList.contains('hidden'));
  }

  function closePreDeckScreen() {
    if (preDeckScreen) {
      preDeckScreen.classList.add('hidden');
    }
  }

  if (document.readyState === 'loading') {
    window.addEventListener('DOMContentLoaded', initDeck);
  } else {
    initDeck();
  }
})();
