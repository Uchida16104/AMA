const scheduleValue = 50;

function doGet() {
  const html = HtmlService.createHtmlOutput(generateHTML())
    .setTitle('出席管理アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function generateHTML() {
  const ss = SpreadsheetApp.openById("SPREAD_SHEET_ID");
  const sheet = ss.getSheetByName("スケジュール");
  const data = sheet.getDataRange().getValues(); // シート全体のデータを取得

  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "MM/dd");
  let html = `
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        *{ text-align: center; }
        body { font-family: Arial, sans-serif; padding: 10px; }
        .card { border: 1px solid #ccc; border-radius: 10px; padding: 10px; margin-bottom: 10px; }
        button { padding: 12px 24px; margin: 3%; font-size: 1em; background: #007bff; color: #fff; border: none; border-radius: 5px; cursor: pointer; }
        button:hover { background: #0056b3; }
        button:disabled { background: #999; }
        a { display: block; margin-bottom: 20px; color: #007bff; }
        h1 { font-size: 4em; }
        h2 { font-size: 2.5em; }
        h3 { font-size: 2em; }
      </style>
    </head>
    <body>
      <h1>出席管理アプリ</h1>
      <h2>今日の日付: ${today}</h2>
  `;

  // スケジュールの開始行とステップ数を定義
  const startRow = 2;
  const step = 5;

  for (let i = 0; i < scheduleValue; i++) {
    const dateRowIndex = startRow + (step * i) - 1; // 2, 6, 10... (0-indexed)
    const placeRowIndex = dateRowIndex + 1;
    const scheduleRowIndex = dateRowIndex + 2;     // 3, 7, 11...
    const planRowIndex = dateRowIndex + 3;         // 4, 8, 12...
    const resultRowIndex = dateRowIndex + 4;       // 5, 9, 13...

    // 範囲外チェック
    if (dateRowIndex >= data.length || placeRowIndex >= data.length || scheduleRowIndex >= data.length || planRowIndex >= data.length || resultRowIndex >= data.length) {
      break;
    }

    const dateRow = data[dateRowIndex];
    const placeRow = data[placeRowIndex];
    const scheduleRow = data[scheduleRowIndex];
    const planRow = data[planRowIndex];
    const resultRow = data[resultRowIndex];

    // B列からH列まで日付を探す
    for (let colIndex = 1; colIndex < Math.min(dateRow.length, 8); colIndex++) {
      const dateValue = dateRow[colIndex];
      if (dateValue) {
        const formattedDate = Utilities.formatDate(new Date(dateValue), "Asia/Tokyo", "MM/dd");
        if (formattedDate === today) {
          const place = placeRow[colIndex];
          const scheduleValue = scheduleRow[colIndex]; // 今日の日付に対応する通所予定の値を取得
          const plan = planRow[colIndex];
          const isChecked = resultRow[colIndex] === true; // チェックボックスの状態

          html += `
            <div class="card" style="text-align: center">
              <h2><strong>場所: </strong>${place || "未設定"}</h2><br>
              <h2><strong>通所予定: </strong>${scheduleValue || "なし"}</h2><br>
              <h3><strong>予定: </strong>${plan || "なし"}</h3><br>
              ${isChecked ? '<button disabled>出席済</button>' : `<button onclick="markAttendance(${resultRowIndex + 1}, ${colIndex + 1}, this)">出席</button>`}
            </div>
          `;
          break; // 今日の日付が見つかったので、この人の他の日付はスキップ
        }
      }
    }
  }

  html += `
    <script>
      function markAttendance(row, col, btn) {
        btn.disabled = true;
        btn.textContent = '処理中...';
        google.script.run.withSuccessHandler(function() {
          btn.textContent = '出席済';
        }).markAttendance(row, col);
      }
    </script>
    <button onclick="updateSchedule()">予定更新</button>
    <script>
      function updateSchedule() {
        const btn = event.target;
        btn.disabled = true;
        btn.textContent = '更新中...';
        google.script.run.withSuccessHandler(function() {
          alert('予定更新済');
          btn.textContent = '予定更新済';
          location.reload();
        }).syncCalendarWithSheet();
      }
    </script>
    <h3><a href="https://docs.google.com/spreadsheets/d/SPREAD_SHEET_ID/edit" target="_blank">スプレッドシートを開く</a></h3>
    </body>
    </html>
  `;

  return html;
}

function markAttendance(row, col) {
  const sheet = SpreadsheetApp.openById("SPREAD_SHEET_ID").getSheetByName("スケジュール");
  const cell = sheet.getRange(row, col);
  cell.setValue(true); // チェックボックスをオンにする
}

function syncCalendarWithSheet() {
  const calendar = CalendarApp.getDefaultCalendar();
  const sheet = SpreadsheetApp.openById("SPREAD_SHEET_ID").getSheetByName("スケジュール");
  const data = sheet.getDataRange().getValues();

  const titleBase = "通所予定";
  const timezone = "Asia/Tokyo";

  const startRow = 2;  // 2行目 (インデックス1)
  const step = 5;      // 各ブロックは5行ごと
  const columns = [1, 2, 3, 4, 5, 6, 7]; // B〜H列（0ベース）

  for (let i = 0; i < scheduleValue; i++) {
    const dateRowIndex = startRow + (step * i) - 1;
    const placeRowIndex = dateRowIndex + 1;
    const scheduleRowIndex = dateRowIndex + 2;

    if (dateRowIndex >= data.length || placeRowIndex >= data.length || scheduleRowIndex >= data.length) break;

    const dateRow = data[dateRowIndex];
    const placeRow = data[placeRowIndex];
    const scheduleRow = data[scheduleRowIndex];

    for (let col of columns) {
      const rawDate = dateRow[col];
      const place = placeRow[col]; // スプレッドシートから場所を取得
      const schedule = scheduleRow[col];

      if (!rawDate) continue;

      const date = new Date(rawDate);
      const formattedDate = Utilities.formatDate(date, timezone, "yyyy/MM/dd");

      // 削除対象：「未定」「なし」または空白
      if (!schedule || schedule === "未定" || schedule === "なし") {
        const possibleSchedules = ["午前", "午後", "全日"];
        for (const s of possibleSchedules) {
          const eventTitle = `${titleBase}（${s}） - ${formattedDate}`;
          const events = calendar.getEventsForDay(date).filter(e => e.getTitle() === eventTitle);
          for (const event of events) {
            event.deleteEvent();
          }
        }
        continue;
      }

      // 登録対象：「午前」「午後」「全日」
      if (["午前", "午後", "全日"].includes(schedule)) {
        let startTime = new Date(date);
        let endTime = new Date(date);

        if (schedule === "午前") {
          startTime.setHours(9, 20, 0);
          endTime.setHours(12, 30, 0);
        } else if (schedule === "午後") {
          startTime.setHours(12, 20, 0);
          endTime.setHours(15, 30, 0);
        } else if (schedule === "全日") {
          startTime.setHours(9, 20, 0);
          endTime.setHours(15, 30, 0);
        }

        const location = place || "未設定"; // 場所がなければ "未設定" とする

        const eventTitle = `${titleBase}（${schedule}） - ${formattedDate}`;
        const existingEvents = calendar.getEventsForDay(date).filter(e => e.getTitle().startsWith(`${titleBase}（`) && e.getTitle().endsWith(`- ${formattedDate}`));

        let updated = false;

        for (const event of existingEvents) {
          if (event.getTitle() === eventTitle) {
            event.setTime(startTime, endTime);
            event.setLocation(location);
            event.setDescription("自動更新された通所予定");
            event.setColor(CalendarApp.EventColor.RED);

            const reminders = event.getPopupReminders();
            if (reminders.length === 0) {
              event.addPopupReminder(0);
            }
            updated = true;
          } else {
            event.deleteEvent();
          }
        }

        if (!updated) {
          calendar.createEvent(eventTitle, startTime, endTime, {
            location,
            description: "自動追加された通所予定",
            color: CalendarApp.EventColor.RED,
            sendInvites: false
          }).addPopupReminder(0);
        }
      }
    }
  }
}


