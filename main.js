const ROUTINE = ['pull', '', 'push', 'aerobic', '','leg', '']; // ジムに行くルーチン
const SN_SCHEDULE = "スケジュール";
const SN_HOLIDAYS = "やすみ";
const NUM_WEEKS = 2; // 先n週間分のスケジュールを生成
const HOLIDAYS_OF_WEEK=[4];//休み曜日を設定、0が日曜、6が土曜

//毎週実行
function Weekly(){
  const ss = SpreadsheetApp.getActive();
  //スケジュール作成
  let sd = createSchedule(ss);
  // 実行日以降のスケジュールを全て削除
  deleteSchedules();
  /// スケジュールをカレンダーに登録
  addEventsToCalendar(sd);
  /// スケジュールをスプレッドシートに書き込み
  writeToSpreadsheet(ss,sd);
}

//スケジュールメインルーチン
function createSchedule(ss) {
  const sc = ss.getSheetByName(SN_SCHEDULE);
  const sc_lastrow = sc.getLastRow();
  let sc_datas =  sc.getRange(2, 1, sc_lastrow-1, 3).getValues();
  let last_data = findNearestPastDate(sc_datas);
  console.log(last_data);
  //スケジュール作成
  let this_data = createSdfromLastSd(last_data);
  //console.log(this_data);
  ///HOLIDAYS_OF_WEEKで指定された部分をあける（後ろにずらす）
  let indexlist = findIndexOnDays(this_data,HOLIDAYS_OF_WEEK);
  this_data = shiftDatesFromIndex(this_data,indexlist);
  ///休みリストで指定された部分をあける（後ろにずらす）
  //2週間分に整形
  this_data = deleteElementsAfterDate(this_data);
  ///ROUTINEの中で空の部分の要素番号を抽出、その部分を削除
  console.log(this_data);
  this_data = this_data.filter(item => item[2] !== "");
  console.log(this_data);
  return this_data;
}

//指定曜日の日付リストを探す、日付リストはインデックス番号で返す
function findIndexOnDays(arr, targetDayIndices) {
  const resultIndices = [];
  for (let i = 0; i < arr.length; i++) {
    const currentDate = arr[i][0];
    const currentDayIndex = currentDate.getDay();
    if (targetDayIndices.includes(currentDayIndex)) {
      resultIndices.push(i);
    }
  }
  return resultIndices;
}

//指定したインデックス番号以降の日付をずらす
function shiftDatesFromIndex(arr, startIndices) {
  for (const startIndex of startIndices) {
    for (let j=startIndex; j<(arr.length-startIndex); j++) {
      arr[j][0] = arr[j+1][0];
    }
  }
  return arr;
}

// 指定した日付のn日後の日付を探してそれ以降を削除
function deleteElementsAfterDate(arr, date = new Date()) {
  // 指定した日付のn日後の日付を計算
  const lim_date = new Date(date);
  lim_date.setDate(lim_date.getDate() + (NUM_WEEKS*7));
  // 指定した日付以降の要素を探し、見つかったらそれ以降の要素を削除
  for (let i = 0; i < arr.length; i++) {
    if (arr[i][0] > lim_date) {
      arr.splice(i);
      break;
    }
  }
  return arr;
}


//n週間分のスケジュールを作成
function createSdfromLastSd(ld = undefined,date = new Date()) {
  date.setHours(0, 0, 0, 0);
  const length = (NUM_WEEKS+1)*7;//予備で多めに作る
  const sc_len = ROUTINE.length;
  let startindex;
  let val = [];
  if (ld === null || ld === undefined){
    startindex = 0;
  }else{
    startindex = ld[1]+1;
    //前回のデータが昨日より前かどうかを調べて、昨日より前だった場合は休日を飛ばすのでstartindexを調整する
    if ((date.getTime() - ld[0].getTime())>(1000 * 60 * 60 * 24)){
      while (ROUTINE[startindex] = ''){startindex++;}
    }
  }
  //startindexがスケジュールの個数よりも多くなってしまったとき調整する
  while (startindex >= sc_len){startindex = startindex-sc_len;}
  //スケジュール作成
  for (let i=0; i<length; i++) {
    let new_date=new Date();
    new_date.setDate(date.getDate() + i);
    new_date.setHours(0, 0, 0, 0);
    val[i] = [];
    val[i][0] = new_date;
    val[i][1] = startindex;
    val[i][2] = ROUTINE[startindex];
    startindex++;
    if (startindex === sc_len){startindex = startindex-sc_len;}
  }
  return val; 
}


//直近の日付のデータを抜き出す
function findNearestPastDate(arr) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let closestDate = null; // 最も近い過去の日付を格納する変数
  let closestIndex = -1; // 対応するインデックスを格納する変数
  for (let i = 0; i < arr.length; i++) {
    let currentDate = arr[i][0];
    // currentDate が今日の日付よりも後であれば、次の日付を確認
    if (currentDate >= today) {
      continue;
    }
    // currentDate が今日の日付よりも前で、かつ最も近い日付よりも近い場合
    if (closestDate === null || currentDate > closestDate) {
      closestDate = currentDate;
      closestIndex = i;
    }
  }
  // 最も近い過去の日付に対応する arr[i] を抜き出す
  let arrB = closestIndex !== -1 ? arr[closestIndex] : null;
  return arrB;
}

// スプレッドシートに書き込む
function writeToSpreadsheet(ss,arr) {
  let sc = ss.getSheetByName(SN_SCHEDULE)
  // 書き込む範囲を指定（A2からCまで）
  const range = sc.getRange(2, 1, arr.length, arr[0].length);
  // スプレッドシートの指定した範囲をクリアする
  clearSpreadsheet(sc);
  // 二次元配列をスプレッドシートに書き込み
  range.setValues(arr);
}

// スプレッドシートの指定した範囲をクリアする
function clearSpreadsheet(sheet) {
  // シートの最終行を取得
  const lastRow = sheet.getLastRow();
  // 指定した範囲をクリア
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
}

//配列からカレンダーに登録
function addEventsToCalendar(arr) {
  // Googleカレンダーの取得
  const calendar = CalendarApp.getCalendarById(CALENDER_ID);
  // 二次元配列を順番に処理
  for (let i = 0; i < arr.length; i++) {
    const eventDate = arr[i][0]; // 日付
    const eventName = arr[i][2]; // イベント名
    // カレンダーにイベントを追加
    const event = calendar.createAllDayEvent(eventName, eventDate);
    Logger.log('イベントを追加しました: ' + event.getTitle());
  }
}
//先n週間分のカレンダー予定を消す
function deleteSchedules(date = new Date()) {
  let twoWeeksLater = new Date();
  twoWeeksLater.setDate(date.getDate() + (NUM_WEEKS*7)); 
  let events = CalendarApp.getCalendarById(CALENDER_ID).getEvents(date, twoWeeksLater);
  for (let i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }
}

