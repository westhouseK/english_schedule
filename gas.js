// スプレットシートの情報
const spreadSheet = {
  id: '', // 各自id
  title: 'englishSchedule',　// 任意のタイトル
  open: () => SpreadsheetApp.getActiveSheet()
}

// Gmailの情報
const gmailApp = {
  from: '', // フィルターするメールアドレス
  searchCount: 10 // 検索する数
}

// TimeTreeの情報
class Timetree {

  constructor(obj) {
    this.calendule = {
      id: '', // カレンダーid
      color: '9'
    };
    this.user = '' // ユーザーid
    this.option = {
      method: 'POST',
      contentType: 'application/json'
    };
    this.headers = {
      Accept: 'application/vnd.timetree.v1+json',
      Authorization: 'Bearer ###' // ###にトークンをセット
    };
    this.url = {
      domain: 'https://timetreeapis.com',
      path: '/calendars//events'　// ###にカレンダーidをセット
    };
    this.option.headers = this.headers;

    obj.data.relationships = {
      label: {
        data: {
          id: `${this.calendule.id}, ${this.calendule.color}`,
          type: 'label'
        }
      },
      attendees: {
        data: [
          {
            id: `${this.calendule.id}, ${this.user}`,
            type: 'user'
          }
        ]
      }
    };
    this.option.payload = JSON.stringify(obj);
  }

  getUrl() {
    return this.url.domain+this.url.path;
  }

  getOption() {
    return this.option;
  }
}

function myFunction() {

  spreadSheet.open();
  // Gmailから10件取得
  const threads = getGmail(gmailApp.from, gmailApp.searchCount);
  // Gmailから予約情報を取得
  const reserveMoments = extractReserveInfo(threads);

  // スプレットシートからレコードを取得
  const records = getReserveRecords(spreadSheet.id, spreadSheet.title);

  Object.keys(reserveMoments).forEach(key => {
    // 予定情報がスプレットシートに存在するかを確認する
    const recordExists = records.some(value => value.format('LLL') === key);
    
    if (recordExists) {
      // すでに登録済み
      Logger.log('レコードが存在するため、登録しない')
      return;
    }
    // Time treeに予定情報を登録
    const timeTree = new Timetree(formatTimetreeData(reserveMoments[key]));
    const res = UrlFetchApp.fetch(timeTree.getUrl(), timeTree.getOption());
    Logger.log(`${res.getResponseCode()}: ${res.getContentText()}`);

    // レコード登録
    SpreadSheetsSQL.open(spreadSheet.id, spreadSheet.title).insertRows([formatResistInfo(reserveMoments[key])]);
    Logger.log('レコードが存在しないので、登録する')
  });
}

// Gamilを取得する
function getGmail(from, num) {
  return GmailApp.getMessagesForThreads(GmailApp.search(from, 0, num));
}

// メールから予定を抽出する 同じ日付であれば上書かれていくので注意。
function extractReserveInfo(mails) {
  const reserveMoments = {}; 
  mails.forEach(thread => { 
    thread.forEach(message => {
      const allSentences = message.getPlainBody().split('\n');
      // 予約情報を抽出 Jodie, Ariel Wild...から始める行を抽出
      const reserveInfo = allSentences.filter(sentence => {
        return (sentence.match(/^[A-Z][a-z]+: /i) ||
        sentence.match(/^[A-Z][a-z]+ [A-Z][a-z]+: /i) ||
        sentence.match(/^[A-Z][a-z]+ [A-Z]: /i))
      }) || [];

      // 一致するものがないまたは「,」が含まれていない
      if (reserveInfo.length === 0 || reserveInfo[0].indexOf(',') === -1) return;

      const moment = createReserveInfo(reserveInfo[0].replace(':', ',').split(',').map(str => str.trim()));
      reserveMoments[moment.moment.format('LLL')] = moment;
    });
  });
  return reserveMoments;
}

// スプレットシートから予約レコードをmomentにして、配列に詰める
function getReserveRecords(id, title) {
  const records = [];
  const reservedRecords = SpreadSheetsSQL.open(id, title).select(['year', 'month', 'date', 'start_time']).result();
  reservedRecords.forEach(record => {
    record.start_time = Moment.moment(record.start_time).format('HH:mm'); // 時間を文字列として取得するため
    records.push(Moment.moment(`${record.year}-${record.month}-${record.date} ${record.start_time}`, 'YYYY-MM-DD HH:mm'));
  });
  return records;
}

// 予約情報をフォーマットする ex. [Jodie, Friday, October 09, 2020, 07:30PM (Asia/Tokyo), 30 minutes]
function createReserveInfo([teacher, day, date, year, start_time, lecture_time]) {

  const moment = Moment.moment(`${date}, ${year}, ${start_time.slice(0, 7)}`, 'LLL');
  return {
    moment,
    teacher,
    lecture_time: parseInt(lecture_time.substr(0, 2))
  }
}

// 予約情報をフォーマットする 
function formatResistInfo({moment, teacher, lecture_time}) {

  return {
    create_date: Moment.moment().format("YYYYMMDDHHmmss"),
    year: moment.format('YYYY'),
    month: moment.format('MM'),
    date: moment.format('DD'),
    day: moment.format('dddd'),
    start_time: moment.format('HH:mm'),
    end_time: moment.clone().add(lecture_time, 'm').format('HH:mm'), // monentが参照渡しのため、cloneする
    lecture_time,
    teacher
  }
}

// time treeに登録するようにデータをフォーマットする
function formatTimetreeData({moment, teacher, lecture_time}) {

  return {
    data: {
      attributes: {
        category: 'schedule',
        title: 'Cambly',
        all_day: false,
        start_at: moment.toISOString(), // ISO8601
        start_timezone: 'UTC',
        end_at: moment.clone().add(lecture_time, 'm').toISOString(),
        end_timezone: 'UTC',
        description: `先生は${teacher}さん`
      }
    }
  }
}


