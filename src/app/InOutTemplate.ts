import * as XLSX from 'xlsx';
import { autofitColumns } from './OTTemplate';
const dummy = {
  no: 1,
  name: 'name',
  date: new Date(),
  in: '07:00',
  out: '08:00',
};

const dummyArray = new Array(100).fill(dummy);
export const toInOutTemplate = () => {
  //
  const emptyRows = Array(3).fill({});
  const rows = dummyArray;
  const header = ['Ma nv', 'Ten NV', 'Ngay', 'Vao', 'Ra', 'Note'];
  const worksheet = XLSX.utils.json_to_sheet([...emptyRows, ...rows], {
    skipHeader: true,
  });
  const merge = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 6} },
  ];
  XLSX.utils.sheet_add_aoa(worksheet, [['CHI TIET ...']], { origin: 'A1' });
  XLSX.utils.sheet_add_aoa(worksheet, [['Tu ngay A den ngay B']], { origin: 'A2' });
  XLSX.utils.sheet_add_aoa(worksheet, [header], { origin: 'A3' });
  worksheet['!merges'] = merge;

  //
  autofitColumns(rows, worksheet, header);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Dates');
  XLSX.writeFile(workbook, 'Presidents.xlsx');
};

