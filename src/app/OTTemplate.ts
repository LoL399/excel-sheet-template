import * as moment from 'moment';
import * as XLSX from 'xlsx';

const dummy = {
  id: Math.random().toString(),
  name: 'To Vinh Loi',
};
const dummyArray = Array(100).fill(dummy);
const dummDate = {
  sdate: moment().set({ date: 1, month: 6 }).toDate(),
  content: 3.5,
};

export const toOTTemplate = () => {
  const month = 6;
  const startDate = moment().set({ date: 26, month: month - 1 });
  const endDate = moment().set({ date: 25, month: month });
  const dateHeader = createObjDate(startDate, endDate);
  const header = [
    'SDT',
    'Ma NV',
    'Ten',
    ...dateHeader.dateHeader,
    'type1',
    'type2',
  ];
  // create rows
  let rows: any[] = [];
  let emptyRows = new Array(2).fill({});
  dummyArray.forEach((element, idx) => {
    let dateObj = {
      no: idx + 1,
      ...element,
      ...dateHeader.dateContent,
      type1: 'test',
      type2: 'test',
    };
    dateObj[moment(dummDate.sdate).format('DD/MM')] = dummDate.content;
    rows.push(dateObj);
  });
  //
  const sumRow = sumObjDate(startDate, endDate, rows);

  const worksheet = XLSX.utils.json_to_sheet([...emptyRows, ...rows], {
    skipHeader: true,
  });
  // decorate
  // create header
  XLSX.utils.sheet_add_aoa(worksheet, [header], { origin: 'A2' });
  // resize auto fit
  // merge header
  if (dateHeader.changeMonth > 0) {
    let startIndex = 3 + dateHeader.changeMonth;
    const merge = [
      { s: { r: 0, c: 3 }, e: { r: 0, c: startIndex - 1 } },
      { s: { r: 0, c: startIndex }, e: { r: 0, c: startIndex + 25 - 1 } },
    ];
    worksheet['!merges'] = merge;
    const prev = startDate.clone().format('MM/YYYY');
    let startPrev = String.fromCharCode(97 + 3).toUpperCase();
    XLSX.utils.sheet_add_aoa(worksheet, [[prev]], {
      origin: `${startPrev}1`,
    });
    //
    const curr = endDate.clone().format('MM/YYYY');
    let currPrev = String.fromCharCode(97 + startIndex).toUpperCase();
    XLSX.utils.sheet_add_aoa(worksheet, [[curr]], { origin: `${currPrev}1` });
    XLSX.utils.sheet_add_aoa(worksheet, [sumRow], {
      origin: `${startPrev}${emptyRows.length + rows.length + 1}`,
    });
  }
  // sum value

  // export file
  autofitColumns(rows, worksheet, header);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Dates');
  XLSX.writeFile(workbook, 'Presidents.xlsx');
};

export const autofitColumns = (
  json: any[],
  worksheet: XLSX.WorkSheet,
  header: string[] = []
) => {
  const jsonKeys = Object.keys(json[0]);
  let objectMaxLength: any[] = [];
  for (let i = 0; i < json.length; i++) {
    let value = json[i];
    for (let j = 0; j < jsonKeys.length; j++) {
      const l = value[jsonKeys[j]] ? value[jsonKeys[j]]?.toString().length : 0;
      objectMaxLength[j] = objectMaxLength[j] >= l ? objectMaxLength[j] : l;
    }

    let key = jsonKeys;
    for (let j = 0; j < key.length; j++) {
      objectMaxLength[j] =
        objectMaxLength[j] >= key[j].length
          ? objectMaxLength[j]
          : key[j].length;
    }
  }

  const wscols = objectMaxLength.map((w) => {
    return { width: w };
  });

  worksheet['!cols'] = wscols;
};

function createObjDate(startDate: moment.Moment, endDate: moment.Moment) {
  let obj = new DateHeader();
  let tickDate = moment(startDate).clone();
  let diffCheck = 0;
  let { dateContent, dateHeader, changeMonth } = obj;
  while (endDate.diff(tickDate, 'days') >= 0) {
    dateContent[tickDate.format('DD/MM')] = '';
    dateHeader.push(tickDate.date().toString());
    if (tickDate.month() !== endDate.month()) {
      diffCheck++;
    }
    tickDate.add(1, 'day');
  }
  obj.changeMonth = diffCheck;
  return obj;
}

// suym by date
const sumObjDate = (
  startDate: moment.Moment,
  endDate: moment.Moment,
  rows: any[]
) => {
  let obj: any[] = [];
  let tickDate = moment(startDate).clone();
  while (endDate.diff(tickDate, 'days') >= 0) {
    let key = tickDate.format('DD/MM');
    let sum = 0;
    rows.forEach((row) => {
      sum += row[key] || 0;
    });
    if (sum === 0) {
      obj.push('');
    } else {
      obj.push(sum);
    }
    tickDate.add(1, 'day');
  }
  return obj;
};

class DateHeader {
  dateContent: any = {};
  dateHeader: string[] = [];
  changeMonth: number = 0;
}
