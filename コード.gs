// ===== 設定 =====
const OPENROUTER_API_KEY = 'ここにAPIキー';
const SPREADSHEET_ID = 'ここにスプレッドシートのID';

// ===== Webアプリのエントリーポイント =====
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('筋トレ記録 自動入力')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== メイン処理 =====
function processImage(base64Image) {
  try {
    const jsonData = analyzeImageWithOpenRouter(base64Image);
    const expanded = expandRecords(jsonData.records);
    const result = writeToSpreadsheet(jsonData.date, expanded);
    return { success: true, message: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ===== レコード展開（カンマ区切り・L〇を行分割） =====
function expandRecords(records) {
  const expanded = [];

  for (const r of records) {
    const menu   = r.menu   ?? '';
    const weight = r.weight ?? '';
    const sets   = r.sets   ?? '';
    const reps   = String(r.reps ?? '');

    const repParts = reps.split(',').map(s => s.trim()).filter(s => s !== '');

    if (repParts.length === 0) {
      expanded.push({ menu, weight, reps: '', sets });
      continue;
    }

    const isSplit = repParts.length > 1;

    for (const part of repParts) {
      const lMatch = part.match(/^[Ll](\d+)$/);

      if (lMatch) {
        // L表記 → セット数は必ず1
        expanded.push({ menu, weight, reps: Number(lMatch[1]), sets: 1 });
      } else if (isSplit) {
        // カンマ分割された行 → セット数は必ず1
        expanded.push({ menu, weight, reps: Number(part), sets: 1 });
      } else {
        // 分割なし → 元のセット数を引き継ぐ
        expanded.push({ menu, weight, reps: Number(part), sets });
      }
    }
  }

  return expanded;
}

// ===== OpenRouter API呼び出し =====
function analyzeImageWithOpenRouter(base64Image) {
  const url = 'https://openrouter.ai/api/v1/chat/completions';

  const prompt = `この手書きの筋トレメモ画像を解析してください。

【ルール】
1. 日付はメニュー名欄の最初の行に書いてあります（例：「6/6 DP」）
   日付のフォーマットは必ず「YY/M/D」形式にしてください（例：「6/6」→「26/6/6」）
   年は現在の西暦下2桁（2026年なら26）を使用してください
2. 回数がカンマ区切り（例：8, 5, 5）の場合はそのままカンマ区切りで記録してください
3. 備考欄に「L5」「L10」などがある場合は、repsフィールドにカンマ区切りで追記してください（例：reps:"6,L5"）
4. weightが空欄の場合は null にしてください
5. setsが空欄の場合は null にしてください
6. 重量は行ごとに注意深く読み取ってください。上の行と異なる数字が書かれている場合は必ず別の重量として記録してください。
7. カンマ区切りに見えても、実際には1つの数字の可能性があります。数字の形をよく確認してください。
8. 備考欄に「L5」「L10」など複数の記載がある場合は、すべて漏らさずrepsにカンマ区切りで追記してください。
9. 同じメニュー名で重量が異なる行が複数ある場合（例：13.75と11.25）は、それぞれ必ず別の行として記録してください。重量の行を絶対にスキップしないでください。

【出力形式】必ず以下のJSON形式のみで返してください。前後に文章・記号・コードブロック一切不要。
{
  "date": "26/6/6",
  "records": [
    {"menu": "DP", "weight": 18, "reps": "8", "sets": 1},
    {"menu": "DP", "weight": 18, "reps": "5", "sets": 2},
    {"menu": "IDP", "weight": 16, "reps": "6,L5", "sets": 2}
  ]
}`;

  const payload = {
    model: 'openrouter/free',
    messages: [
      {
        role: 'system',
        content: 'You are a JSON-only API. Output only valid JSON starting with { and ending with }. No markdown, no explanation, no code blocks.'
      },
      {
        role: 'user',
        content: [
          {
            type: 'image_url',
            image_url: { url: `data:image/jpeg;base64,${base64Image}` }
          },
          {
            type: 'text',
            text: prompt
          }
        ]
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${OPENROUTER_API_KEY}`,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'Workout Tracker'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseJson = JSON.parse(response.getContentText());

  if (responseJson.error) {
    throw new Error(responseJson.error.message);
  }

  if (!responseJson.choices || !responseJson.choices[0]) {
    throw new Error('APIレスポンス異常: ' + response.getContentText());
  }

  const content = responseJson.choices[0].message.content;
  Logger.log('モデル応答: ' + content);

  // contentがnullまたは空の場合（画像非対応モデルが選ばれた場合）
  if (!content || content.trim() === '') {
    throw new Error('モデルが画像を処理できませんでした。もう一度試してください。');
  }

  const jsonMatch = content.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('JSONが見つかりません。返答: ' + content.substring(0, 200));
  }

  return JSON.parse(jsonMatch[0]);
}

// ===== スプレッドシートへの書き込み =====
function writeToSpreadsheet(date, records) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const existing = ss.getSheetByName(date);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(date, 0);

  // ヘッダー（B〜E列、備考列なし）
  sheet.getRange('B2:E2').setValues([['メニュー名', '重量', '回数', 'セット数']]);
  sheet.getRange('B2:E2').setFontWeight('bold');

  // データ書き込み
  for (let i = 0; i < records.length; i++) {
    const row = i + 2;
    const r = records[i];
    sheet.getRange(row, 2).setValue(r.menu   ?? '');
    sheet.getRange(row, 3).setValue(r.weight ?? '');
    sheet.getRange(row, 4).setValue(r.reps   ?? '');
    sheet.getRange(row, 5).setValue(r.sets   ?? '');
  }

  // プルダウンをデータ行のみに設定
  const menuSheet = ss.getSheetByName('メニューリスト');
  if (menuSheet && records.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(menuSheet.getRange('B2:B'), true)
      .build();
    sheet.getRange(2, 2, records.length, 1).setDataValidation(rule);
  }

  return `「${date}」シートを作成し、${records.length}件のデータを入力しました！`;
}

function openWebApp() {
  const url = 'https://script.google.com/macros/s/AKfycbyCScGoS-kjtdinTDPvQDWZU_P3UYRiP8ReQTHA85NvAk32rUe0_t1ry6yWHxkcSAt-/exec';
  const html = `<script>window.open('${url}','_blank');google.script.host.close();</script>`;
  const ui = HtmlService.createHtmlOutput(html).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Webアプリを開いています...');
}

// ===== チェックボックス連動トリガー =====
function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // メニューリストシートのみ対応
  if (sheetName !== 'メニューリスト') return;
  
  // チェックボックスのセル位置（例：J2）を監視
  const targetCell = 'J2';
  if (e.range.getA1Notation() !== targetCell) return;
  
  // チェックがONになった時だけ動く
  if (e.value !== 'TRUE') return;

  // チェックを自動でOFFに戻す（次回も使えるように）
  sheet.getRange(targetCell).setValue(false);

  // WebアプリのURLを開く
  const webAppUrl = 'https://script.google.com/macros/s/AKfycbyCScGoS-kjtdinTDPvQDWZU_P3UYRiP8ReQTHA85NvAk32rUe0_t1ry6yWHxkcSAt-/exec';
  const html = `
    <html>
      <body>
        <script>
          window.open('${webAppUrl}', '_blank');
          google.script.host.close();
        </script>
      </body>
    </html>`;
  const ui = HtmlService.createHtmlOutput(html).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Webアプリを開いています...');
}

// ===== デバッグ用 =====
function debugOpenRouter() {
  const url = 'https://openrouter.ai/api/v1/chat/completions';

  const payload = {
    model: 'openrouter/free',
    messages: [
      { role: 'user', content: 'こんにちは。一言で返してください。' }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${OPENROUTER_API_KEY}`,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'Workout Tracker'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log('ステータス: ' + response.getResponseCode());
  Logger.log('レスポンス: ' + response.getContentText());
}

function debugAllModels() {
  const models = [
    'google/gemma-4-31b-it:free',
    'google/gemma-4-26b-a4b-it:free',
    'openrouter/free'
  ];

  const url = 'https://openrouter.ai/api/v1/chat/completions';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${OPENROUTER_API_KEY}`,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'Workout Tracker'
    },
    muteHttpExceptions: true
  };

  for (const model of models) {
    const payload = {
      model: model,
      messages: [{ role: 'user', content: 'テスト' }]
    };
    options.payload = JSON.stringify(payload);
    const response = UrlFetchApp.fetch(url, options);
    Logger.log(model + ' → ' + response.getResponseCode() + ' : ' + response.getContentText().substring(0, 100));
  }
}
