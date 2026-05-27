// ============================================================
// Priming Experiment — Google Apps Script 後端
// 部署方式：擴充功能 > Apps Script > 部署 > 新增部署
// 類型選「網頁應用程式」，存取權限設「所有人」
// ============================================================

const SHEET_ID = '1vsjV7N1OJrj2Gav7eNfQbfY1haXFQWkJlrkmm-4JiUE'; // <-- 填入你的 Sheet ID

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getConfig') return getConfig();
  if (action === 'getTrials') return getTrials();
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'saveResults') return saveResults(data);
    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ── 讀取 config 工作表 ──────────────────────────────────────
function getConfig() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('config');
  const data = sheet.getDataRange().getValues();

  const config = {};
  // config 工作表格式：A欄 = key, B欄 = value
  // key 包含：welcome_title, welcome_body, consent_text,
  //           survey_fields (JSON array), instruction_text,
  //           practice_instruction, experiment_instruction,
  //           iti_ms, prime_ms, response_timeout_ms
  data.forEach(row => {
    if (row[0]) config[row[0]] = row[1];
  });

  // survey_fields 是 JSON 字串，需解析
  if (config.survey_fields) {
    try { config.survey_fields = JSON.parse(config.survey_fields); }
    catch (e) { config.survey_fields = []; }
  }

  return jsonResponse({ ok: true, config });
}

// ── 讀取 trials 工作表 ──────────────────────────────────────
function getTrials() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('trials');
  const [header, ...rows] = sheet.getDataRange().getValues();

  // trials 工作表欄位：
  // type | item_id | prime_a | prime_b | target_audio_url | answer | notes
  const trials = rows
    .filter(r => r[0]) // 過濾空行
    .map(r => ({
      type:           r[0], // 'practice' 或 'experiment'
      item_id:        r[1],
      prime_a:        r[2],
      prime_b:        r[3],
      target_url:     r[4],
      answer:         String(r[5]).toUpperCase() === 'TRUE',
      notes:          r[6] || ''
    }));

  return jsonResponse({ ok: true, trials });
}

// ── 儲存實驗結果 ─────────────────────────────────────────────
function saveResults(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('results');
  if (!sheet) {
    sheet = ss.insertSheet('results');
    // 建立標題列
    sheet.appendRow([
      'session_id', 'timestamp', 'participant_id',
      // 受試者資料（動態欄位）
      'survey_data',
      // 裝置環境
      'is_mobile', 'os', 'browser', 'audio_base_latency_ms',
      'hardware_concurrency', 'user_agent',
      // 每筆試次
      'trial_index', 'type', 'item_id',
      'prime_assigned', 'prime_text',
      'target_url', 'correct_answer',
      'response', 'accuracy', 'rt_ms',
      'timed_out', 'notes'
    ]);
  }

  const { session_id, participant_id, survey_data,
          device_info, trial_results } = data;
  const timestamp = new Date().toISOString();

  trial_results.forEach(trial => {
    sheet.appendRow([
      session_id,
      timestamp,
      participant_id,
      JSON.stringify(survey_data),
      device_info.is_mobile,
      device_info.os,
      device_info.browser,
      device_info.audio_base_latency_ms,
      device_info.hardware_concurrency,
      device_info.user_agent,
      trial.trial_index,
      trial.type,
      trial.item_id,
      trial.prime_assigned,   // 'A' 或 'B'
      trial.prime_text,
      trial.target_url,
      trial.correct_answer,
      trial.response,         // true / false / null(timeout)
      trial.accuracy,         // 1 / 0 / null
      trial.rt_ms,
      trial.timed_out,
      trial.notes
    ]);
  });

  return jsonResponse({ ok: true, saved: trial_results.length });
}

// ── 工具函式 ─────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}