// ============================================================
// 執行此腳本可自動建立 Google Sheet 範本結構
// 在 Apps Script 編輯器中執行 setupSheet() 函式即可
// ============================================================

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 建立 config 工作表 ──────────────────────────────────
  let config = ss.getSheetByName('config');
  if (!config) config = ss.insertSheet('config');
  else config.clearContents();

  config.getRange('A1:B1').setValues([['key', 'value']]);
  config.getRange('A1:B1').setFontWeight('bold');

  const configData = [
    ['welcome_title', '台語詞彙辨識實驗'],
    ['welcome_body', '感謝您參與本研究。\n本實驗旨在探討華語詞彙對台語詞彙辨識的影響。\n全程約需 15-20 分鐘，請在安靜的環境中使用耳機進行。'],
    ['consent_text', '本研究已通過倫理審查。您的資料將以匿名方式處理，僅供學術研究使用。\n參與完全自願，您可隨時退出。'],
    ['survey_fields', JSON.stringify([
      {"key":"age","label":"年齡","type":"number","required":true},
      {"key":"gender","label":"性別","type":"radio","options":["男","女","非二元性別","不願透露"],"required":true},
      {"key":"taiwanese_level","label":"台語能力自評（1=完全不會，5=母語程度）","type":"radio","options":["1","2","3","4","5"],"required":true},
      {"key":"mandarin_level","label":"華語能力自評（1=初學，5=母語程度）","type":"radio","options":["1","2","3","4","5"],"required":true},
      {"key":"bilingual_age","label":"幾歲開始接觸台語（若無請填 0）","type":"number","required":false},
      {"key":"education","label":"最高學歷","type":"select","options":["高中以下","高中/職","大學/專科","碩士","博士"],"required":true}
    ])],
    ['practice_instruction', '接下來是幾題練習題，幫助您熟悉流程。\n\n每題流程如下：\n1. 畫面出現「+」注視點\n2. 出現一個華語詞彙（Prime），快速閱讀即可\n3. 聽到一個台語詞彙的發音\n4. 判斷您聽到的是否為真實的台語詞彙\n\n按 F 鍵（或點「詞彙」）= 真實台語詞彙\n按 J 鍵（或點「非詞」）= 不是台語詞彙\n\n練習階段會顯示答案對錯，請放輕鬆。'],
    ['experiment_instruction', '練習完成！接下來是正式實驗。\n\n流程與練習相同，但不會顯示答案對錯。\n\n請盡量：\n・快速作答（越快越好）\n・準確作答（判斷是否為真實台語詞彙）\n\n若不確定，以直覺判斷即可。'],
    ['result_message', '實驗已全部完成，非常感謝您的參與！\n如有任何問題，請聯繫研究人員。'],
    ['iti_ms', '800'],
    ['prime_ms', '300'],
    ['response_timeout_ms', '2000'],
  ];

  config.getRange(2, 1, configData.length, 2).setValues(configData);
  config.setColumnWidth(1, 200);
  config.setColumnWidth(2, 600);
  config.getRange('B:B').setWrap(true);

  // ── 建立 trials 工作表 ──────────────────────────────────
  let trials = ss.getSheetByName('trials');
  if (!trials) trials = ss.insertSheet('trials');
  else trials.clearContents();

  const trialsHeader = ['type', 'item_id', 'prime_a', 'prime_b', 'target_audio_url', 'answer', 'notes'];
  trials.getRange(1, 1, 1, trialsHeader.length).setValues([trialsHeader]);
  trials.getRange(1, 1, 1, trialsHeader.length).setFontWeight('bold');

  // 範例資料（練習題 + 正式題）
  const trialsData = [
    // 練習題（type = practice）
    ['practice', 'P01', '客老', '火乾', 'https://example.com/audio/P01.mp3', 'TRUE',  '練習：洘流 khò-lâu（真詞）'],
    ['practice', 'P02', '白色', '紅色', 'https://example.com/audio/P02.mp3', 'FALSE', '練習：假詞（非詞）'],
    ['practice', 'P03', '海風', '山雨', 'https://example.com/audio/P03.mp3', 'TRUE',  '練習：真詞'],
    // 正式題（type = experiment）
    ['experiment', 'E01', '客老', '火乾', 'https://example.com/audio/E01.mp3', 'TRUE',  '目標：洘流 khò-lâu'],
    ['experiment', 'E02', '時間', '空氣', 'https://example.com/audio/E02.mp3', 'FALSE', '目標：非詞'],
    ['experiment', 'E03', '大聲', '小聲', 'https://example.com/audio/E03.mp3', 'TRUE',  '目標：台語真詞'],
    ['experiment', 'E04', '花香', '草香', 'https://example.com/audio/E04.mp3', 'FALSE', '目標：非詞'],
  ];

  trials.getRange(2, 1, trialsData.length, trialsHeader.length).setValues(trialsData);

  // 欄寬設定
  [1,2,3,4,6,7].forEach((col, i) => trials.setColumnWidth(col, [120,80,100,100,80,200][i]));
  trials.setColumnWidth(5, 350);

  // answer 欄位下拉選單
  const answerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true).build();
  trials.getRange('F2:F1000').setDataValidation(answerRule);

  // type 欄位下拉選單
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['practice', 'experiment'], true).build();
  trials.getRange('A2:A1000').setDataValidation(typeRule);

  // ── 建立 results 工作表（空白，由 GAS 自動填入）──────
  let results = ss.getSheetByName('results');
  if (!results) results = ss.insertSheet('results');
  else results.clearContents();

  const resultsHeader = [
    'session_id', 'timestamp', 'participant_id', 'survey_data',
    'is_mobile', 'os', 'browser', 'audio_base_latency_ms', 'hardware_concurrency', 'user_agent',
    'trial_index', 'type', 'item_id', 'prime_assigned', 'prime_text',
    'target_url', 'correct_answer', 'response', 'accuracy', 'rt_ms',
    'timed_out', 'notes'
  ];
  results.getRange(1, 1, 1, resultsHeader.length).setValues([resultsHeader]);
  results.getRange(1, 1, 1, resultsHeader.length).setFontWeight('bold');
  results.setFrozenRows(1);

  // ── 提示 ────────────────────────────────────────────────
  SpreadsheetApp.getUi().alert(
    '✅ 設定完成！\n\n' +
    '已建立以下工作表：\n' +
    '・config  — 實驗設定\n' +
    '・trials  — 題目（含範例資料）\n' +
    '・results — 自動儲存結果\n\n' +
    '接下來請：\n' +
    '1. 修改 config 工作表中的文字與時間設定\n' +
    '2. 在 trials 工作表填入您的題目\n' +
    '3. 在 Code.gs 中填入本試算表的 ID\n' +
    '4. 部署為 Web App'
  );
}