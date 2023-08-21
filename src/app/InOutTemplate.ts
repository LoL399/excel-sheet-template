import * as XLSX from 'xlsx-js-style';
import { autofitColumns } from './OTTemplate';
import * as moment from 'moment';
const dummy = {
  no: 1,
  name: 'name',
  date: moment().format(),
  in: '07:00',
  out: '08:00',
};

const dummyArray = new Array(100).fill(dummy);
export const toInOutTemplate = () => {
  //
  let rows = [];
  rows = [
    [{ v: "CHI TIET ...", t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } }],
    [{ v: "Tu ngay A den ngay B", t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } }],
    [
      { v: "Ma nv", t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } },
      { v: "Ten NV", t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } },
      { v: "Ngay", t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } },
      { v: "Vao", t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } },
      { v: "Ra", t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } },
      { v: "Note", t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } },
    ],
    ...dummyArray.map((row)=>{
      let key = Object.keys(row);
      return [key.map(k=> {return { v: row[k], t: "s", s: { alignment: { vertical: 'center', horizontal: 'center' }, font: {bold: true} } }})]
    }).flat()
  ]
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const merge = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 6} },
  ];
  ws['!merges'] = merge;

  //
  // autofitColumns(rows, worksheet, header);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, ws, 'Dates');
  XLSX.writeFile(workbook, 'Presidents.xlsx');
};
