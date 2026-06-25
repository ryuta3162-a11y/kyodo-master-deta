function createJoyfitScheduleFixed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "経堂PG";
  
  // シート初期化
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet();
  sheet.setName(sheetName);

  // --- 設定 ---
  const startHour = 8;
  const endHour = 23;
  const days = ["月", "火", "水", "木", "金", "土", "日"];
  const fontFamily = "Roboto"; 

  // 色定義
  const colors = {
    heat: "#E91E63",     // ヒート
    hot: "#F57C00",      // ホット
    yuru: "#FBC02D",     // ゆるホット
    soft: "#00ACC1",     // ソフト
    
    dateBg: "#D32F2F",   // 日付背景（赤）
    mainDark: "#006064", // 曜日ヘッダー
    timeBg: "#00838F",   // 時間軸背景
    timeText: "#FFFFFF", 
    titleText: "#000000",
    starColor: "#FFD700", // 星（ゴールド）
    
    hotLabelBg: "#FF7043",
    pilatesLabelBg: "#42A5F5"
  };

  // --- データ ---
  const scheduleData = [
    {day: "月", area: "ホットスタジオ", start: "09:30", end: "10:30", title: "朝ヨガ", star: "★", instructor: "kazu", type: "yuru"},
    {day: "月", area: "ピラティスルーム", start: "10:30", end: "11:30", title: "ピラティスリフォーマー\nアーム＆ショルダー", star: "★★★", instructor: "YUKI", type: "soft"},
    {day: "月", area: "ホットスタジオ", start: "10:45", end: "11:45", title: "フローヨガ", star: "★★", instructor: "kazu", type: "yuru"},
    {day: "月", area: "ピラティスルーム", start: "12:00", end: "13:00", title: "ピラティスリフォーマー\nチェストオープン", star: "★★", instructor: "YUKI", type: "soft"},
    {day: "月", area: "ホットスタジオ", start: "12:15", end: "13:15", title: "全身ほぐしヨガ", star: "★", instructor: "EMI", type: "yuru"},
    {day: "月", area: "ホットスタジオ", start: "13:45", end: "14:45", title: "美しい骨盤 /\nリンパリフレッシュヨガ", star: "★", instructor: "MIE", type: "hot", note: "隔週"},
    {day: "月", area: "ホットスタジオ", start: "15:15", end: "16:15", title: "かんたんサーキット", star: "★★★", instructor: "MIE", type: "yuru"},
    {day: "月", area: "ホットスタジオ", start: "18:30", end: "19:15", title: "リラックスヨガ", star: "★", instructor: "坂東", type: "hot"},
    {day: "月", area: "ホットスタジオ", start: "19:30", end: "20:30", title: "カラダ調整ヨガ", star: "★★", instructor: "坂東", type: "hot"},
    {day: "火", area: "ホットスタジオ", start: "09:30", end: "10:30", title: "ハタヨガ", star: "★～★★", instructor: "ナカシマ トオル", type: "hot"},
    {day: "火", area: "ホットスタジオ", start: "10:45", end: "11:45", title: "パワーヨガ", star: "★★～★★★", instructor: "ナカシマ トオル", type: "hot"},
    {day: "火", area: "ホットスタジオ", start: "12:15", end: "13:15", title: "おやすみヨガ", star: "★", instructor: "YUKI", type: "hot"},
    {day: "火", area: "ホットスタジオ", start: "13:45", end: "14:45", title: "アグニヨガ", star: "★★★", instructor: "YUKI", type: "yuru"},
    {day: "火", area: "ピラティスルーム", start: "18:15", end: "19:15", title: "ピラティスリフォーマー\nアーム＆ショルダー", star: "★★★", instructor: "みと", type: "soft"},
    {day: "火", area: "ホットスタジオ", start: "19:30", end: "20:45", title: "ポール＆ヨガベーシック", star: "★", instructor: "Shikano", type: "yuru"},
    {day: "火", area: "ピラティスルーム", start: "19:45", end: "20:45", title: "ピラティスリフォーマー\nウエストシェイプ", star: "★★★", instructor: "みと", type: "soft"},
    {day: "水", area: "ホットスタジオ", start: "09:30", end: "10:30", title: "朝ヨガ", star: "★", instructor: "Mariko", type: "hot"},
    {day: "水", area: "ホットスタジオ", start: "10:45", end: "11:45", title: "シェイプアップヨガ", star: "★★", instructor: "Mariko", type: "hot"},
    {day: "水", area: "ホットスタジオ", start: "13:30", end: "14:30", title: "リズムステップ", star: "★★", instructor: "HACHI", type: "yuru"},
    {day: "水", area: "ホットスタジオ", start: "15:00", end: "16:00", title: "美脚美尻ヨガ", star: "★★", instructor: "HACHI", type: "hot"},
    {day: "水", area: "ホットスタジオ", start: "18:00", end: "19:00", title: "ホイールで月礼拝", star: "★★", instructor: "Hana", type: "yuru"},
    {day: "水", area: "ピラティスルーム", start: "18:15", end: "19:15", title: "Core Functional Yoga /\n月礼拝ヨガ", star: "★★★/★★", instructor: "みと", type: "hot", note: "隔週"},
    {day: "水", area: "ホットスタジオ", start: "19:30", end: "20:30", title: "整えるヨガ", star: "★★", instructor: "Hana", type: "hot"},
    {day: "木", area: "ホットスタジオ", start: "10:00", end: "11:00", title: "チャクラヒーリングヨガ", star: "★★", instructor: "MIE", type: "yuru"},
    {day: "木", area: "ホットスタジオ", start: "11:30", end: "12:30", title: "腸活ヨガ /\n美しい骨盤", star: "★", instructor: "MIE", type: "hot", note: "隔週"},
    {day: "木", area: "ピラティスルーム", start: "12:45", end: "13:45", title: "ピラティスリフォーマー\n魅せるバック＆アームズ", star: "★★★", instructor: "Hana", type: "soft"},
    {day: "木", area: "ホットスタジオ", start: "13:00", end: "14:00", title: "スパインコントロールヨガ", star: "★★", instructor: "YUKI", type: "hot"},
    {day: "木", area: "ピラティスルーム", start: "14:15", end: "15:15", title: "ピラティスリフォーマー\n桃尻メイク", star: "★★★", instructor: "Hana", type: "soft"},
    {day: "木", area: "ホットスタジオ", start: "14:30", end: "15:30", title: "SHiN癒", star: "★", instructor: "YUKI", type: "hot"},
    {day: "木", area: "ホットスタジオ", start: "19:45", end: "20:45", title: "ホットピラティス", star: "★★", instructor: "みと", type: "hot"},
    {day: "金", area: "ホットスタジオ", start: "09:30", end: "10:30", title: "朝ヨガ", star: "★", instructor: "kanako", type: "hot"},
    {day: "金", area: "ホットスタジオ", start: "11:00", end: "12:00", title: "美bodyピラティス", star: "★★", instructor: "後藤 亜也", type: "yuru"},
    {day: "金", area: "ホットスタジオ", start: "12:15", end: "13:15", title: "顔ヨガ", star: "★", instructor: "後藤 亜也", type: "yuru"},
    {day: "金", area: "ホットスタジオ", start: "13:45", end: "14:45", title: "ホイールで月礼拝", star: "★★", instructor: "MIE", type: "yuru"},
    {day: "金", area: "ホットスタジオ", start: "15:15", end: "16:15", title: "SHiN美 /\nCore Functional Yoga", star: "★★★/★★", instructor: "MIE", type: "hot", note: "隔週"},
    {day: "金", area: "ホットスタジオ", start: "19:00", end: "20:15", title: "ハタヨガ", star: "★～★★", instructor: "ナカシマ トオル", type: "hot"},
    {day: "土", area: "ホットスタジオ", start: "09:30", end: "10:30", title: "フローヨガ", star: "★★", instructor: "kanako", type: "hot"},
    {day: "土", area: "ホットスタジオ", start: "11:00", end: "12:00", title: "ベーシックヨガ", star: "★", instructor: "kanako", type: "hot"},
    {day: "土", area: "ホットスタジオ", start: "12:30", end: "13:30", title: "全身巡らせ整えるヨガ /\nアロマリラックス", star: "★★", instructor: "EMI", type: "yuru", note: "隔週"},
    {day: "土", area: "ホットスタジオ", start: "14:15", end: "15:15", title: "パワーヴィンヤサ", star: "★★", instructor: "itsuku", type: "hot"},
    {day: "土", area: "ホットスタジオ", start: "15:45", end: "16:45", title: "体幹とバランスを鍛えるヨガ", star: "★★★", instructor: "itsuku", type: "hot"},
    {day: "日", area: "ホットスタジオ", start: "09:30", end: "10:30", title: "フレッチャーピラティス", star: "★★", instructor: "菊池 智子", type: "yuru"},
    {day: "日", area: "ピラティスルーム", start: "11:00", end: "12:00", title: "ピラティスリフォーマー\n魅せるバック＆アームズ", star: "★★★", instructor: "みと", type: "soft"},
    {day: "日", area: "ホットスタジオ", start: "12:45", end: "13:45", title: "BODY BALANCE", star: "★★★", instructor: "坂東", type: "soft"},
    {day: "日", area: "ホットスタジオ", start: "14:00", end: "14:45", title: "カラダ調整ヨガ", star: "★", instructor: "坂東", type: "hot"},
    {day: "日", area: "ホットスタジオ", start: "15:15", end: "16:30", title: "ポール&ヨガベーシック", star: "★☆", instructor: "Shikano", type: "yuru"},
    {day: "日", area: "ホットスタジオ", start: "16:45", end: "17:45", title: "リラックスほぐし＆骨盤調整", star: "★", instructor: "Shikano", type: "yuru"}
  ];

  // 全体のフォント設定
  sheet.getRange("A:O").setFontFamily(fontFamily);

  // --- 1. 新ヘッダーエリアの構築 ---
  
  // A. 日付エリア (左上: A1-A2結合) ★ここを修正しました
  // 固定列（A列）内だけで完結させます
  const dateRange = sheet.getRange("A1:A2");
  dateRange.merge()
           .setValue("2026年\n2月")
           .setBackground(colors.dateBg)
           .setFontColor("white")
           .setFontSize(14) // 幅に合わせて少し小さく
           .setFontWeight("bold")
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle");

  // B. タイトルエリア (中央: B1-I2結合)
  // タイトルをB列から開始するように変更
  const titleRange = sheet.getRange("B1:I2");
  titleRange.merge()
            .setValue("JOYFIT24経堂\nスタジオプログラムスケジュール")
            .setFontSize(20)
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setFontColor(colors.titleText);

  // C. 凡例ボタンエリア (右側: J1-O2)
  const legendLabels = [
    { text: "ヒート\n38-39℃", color: colors.heat, cols: 1 }, 
    { text: "ホット\n36-38℃", color: colors.hot, cols: 1 },  
    { text: "ゆるホット\n30-33℃", color: colors.yuru, cols: 2 }, 
    { text: "ソフト\n床暖28℃", color: colors.soft, cols: 2 }   
  ];

  let currentLegCol = 10; // J列
  legendLabels.forEach(leg => {
    const legRange = sheet.getRange(1, currentLegCol, 2, leg.cols);
    legRange.merge()
            .setValue(leg.text)
            .setBackground(leg.color)
            .setFontColor("white")
            .setFontSize(9)
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setWrap(true);
            
    legRange.setBorder(true, true, true, true, null, null, "white", SpreadsheetApp.BorderStyle.SOLID_THICK);
    currentLegCol += leg.cols;
  });

  // ヘッダー全体の外枠 (A1:O2)
  sheet.getRange("A1:O2").setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);


  // --- 2. 曜日ヘッダー ---
  const dayHeaderRow = 3;
  sheet.getRange(dayHeaderRow, 1).setValue("TIME").setBackground(colors.timeBg).setFontColor(colors.timeText).setFontWeight("bold").setHorizontalAlignment("center");
  
  days.forEach((day, index) => {
    const colStart = 2 + (index * 2);
    const range = sheet.getRange(dayHeaderRow, colStart, 1, 2);
    range.merge().setValue(day + "曜日")
         .setHorizontalAlignment("center").setVerticalAlignment("middle")
         .setFontWeight("bold").setFontSize(12)
         .setBackground(colors.mainDark).setFontColor("white")
         .setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
  });

  // --- 3. エリアヘッダー ---
  const areaHeaderRow = 4;
  days.forEach((day, index) => {
    const colStart = 2 + (index * 2);
    sheet.getRange(areaHeaderRow, colStart).setValue("ホットスタジオ")
         .setBackground(colors.hotLabelBg).setFontColor("white")
         .setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(areaHeaderRow, colStart + 1).setValue("ピラティスルーム")
         .setBackground(colors.pilatesLabelBg).setFontColor("white")
         .setFontSize(8).setFontWeight("bold").setHorizontalAlignment("center");
  });

  // --- 4. 時間軸 ---
  const interval = 15;
  const startRow = 5;
  let currentRow = startRow;
  
  const totalMinutes = (endHour - startHour) * 60;
  const totalRowsCalc = totalMinutes / interval;
  
  sheet.getRange(startRow, 1, totalRowsCalc, 1).setBackground(colors.timeBg).setFontColor(colors.timeText).setFontWeight("bold");

  for (let h = startHour; h < endHour; h++) {
    for (let m = 0; m < 60; m += interval) {
      if (m === 0) {
        let cell = sheet.getRange(currentRow, 1);
        cell.setValue(`${h}:00 ▶`); 
        cell.setHorizontalAlignment("right").setVerticalAlignment("middle").setFontSize(10);
        cell.setBorder(true, null, null, true, null, null, "white", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(currentRow, 2, 4, 14).setBorder(true, null, null, null, null, null, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);
      }
      currentRow++;
    }
  }
  const totalRows = currentRow - startRow;

  // --- 5. レッスンデータ ---
  scheduleData.forEach(lesson => {
    const dayIndex = days.indexOf(lesson.day);
    if (dayIndex === -1) return;
    
    const areaOffset = lesson.area === "ホットスタジオ" ? 0 : 1;
    const targetCol = 2 + (dayIndex * 2) + areaOffset;

    const startParts = lesson.start.split(":");
    const endParts = lesson.end.split(":");
    const startH = parseInt(startParts[0]);
    const startM = parseInt(startParts[1]);
    const endH = parseInt(endParts[0]);
    const endM = parseInt(endParts[1]);

    const rowOffset = ((startH - startHour) * (60 / interval)) + (startM / interval);
    const lessonStartRow = startRow + rowOffset;
    const durationMin = (endH * 60 + endM) - (startH * 60 + startM);
    const numRows = durationMin / interval;

    const targetRange = sheet.getRange(lessonStartRow, targetCol, numRows, 1);
    targetRange.merge();

    const timeStr = `${lesson.start}-${lesson.end}`;
    const titleStr = lesson.title;
    const starStr = lesson.star;
    const nameStr = lesson.instructor;
    const fullText = `${timeStr}\n\n${titleStr}\n${starStr}  ${nameStr}`;

    const bgColor = colors[lesson.type] || "#CCCCCC";
    const isYellowBg = (lesson.type === "yuru");
    const baseTextColor = isYellowBg ? "black" : "white";
    
    const richTextBuilder = SpreadsheetApp.newRichTextValue();
    richTextBuilder.setText(fullText);

    const timeStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(9).setForegroundColor(baseTextColor).build();
    const titleStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(11).setForegroundColor(baseTextColor).build();
    // 変更点: 星の色をテキストカラー(baseTextColor)と同じに設定
    const starStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(10).setForegroundColor(baseTextColor).build(); 
    const nameStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(10).setForegroundColor(baseTextColor).build();

    let currentIdx = 0;
    const timeEnd = currentIdx + timeStr.length;
    richTextBuilder.setTextStyle(currentIdx, timeEnd, timeStyle);
    currentIdx = timeEnd + 2; 
    const titleEnd = currentIdx + titleStr.length;
    richTextBuilder.setTextStyle(currentIdx, titleEnd, titleStyle);
    currentIdx = titleEnd + 1;
    const starEnd = currentIdx + starStr.length;
    richTextBuilder.setTextStyle(currentIdx, starEnd, starStyle);
    currentIdx = starEnd + 2;
    const nameEnd = currentIdx + nameStr.length;
    richTextBuilder.setTextStyle(currentIdx, nameEnd, nameStyle);

    targetRange.setRichTextValue(richTextBuilder.build());
    targetRange.setBackground(bgColor);
    targetRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    targetRange.setVerticalAlignment("middle");
    targetRange.setHorizontalAlignment("center");
    
    targetRange.setBorder(true, true, true, true, null, null, "white", SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  // --- 6. レイアウト調整 ---
  // A列(日付・時間)を少し広くして、日付が潰れないようにする
  sheet.setColumnWidth(1, 80); 
  for (let i = 2; i <= 15; i++) {
    sheet.setColumnWidth(i, 130);
  }
  
  sheet.setRowHeights(1, 2, 40); 
  sheet.setRowHeights(startRow, totalRows, 25);
  
  // 外枠
  sheet.getRange(dayHeaderRow, 2, totalRows + 2, 14).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // 固定 (ここでエラーが出ないように日付の結合をA列のみにしました)
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);
}