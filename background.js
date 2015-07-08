
console.log("Ataze!");
// function datenum(v, date1904) {
// 	if(date1904) v+=1462;
// 	var epoch = Date.parse(v);
// 	return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
// }
// function sheetFromArrayOfObjects(data, opts) {
//   var ws = {};
//   var range = {
//     s: {
//       c: 10000000,
//       r: 10000000
//     },
//     e: {
//       c: 0,
//       r: 0
//     }
//   };
//   var C;
//   for (var R = 0; R != data.length; ++R) {
//     C = 0;
//     for (var key in data[R]) {
//       if (range.s.r > R) range.s.r = R;
//       if (range.s.c > C) range.s.c = C;
//       if (range.e.r < R) range.e.r = R;
//       if (range.e.c < C) range.e.c = C;
//       var cell = {
//         v: data[R][key]
//       };
//       if (cell.v == null) {
//         cell.v = " ";
//       }
//       var cell_ref = XLSX.utils.encode_cell({
//         c: C,
//         r: R
//       });
//
//       if (typeof cell.v === 'number') cell.t = 'n';
//       else if (typeof cell.v === 'boolean') cell.t = 'b';
//       else if (cell.v instanceof Date) {
//         cell.t = 'n';
//         cell.z = XLSX.SSF._table[14];
//         cell.v = datenum(cell.v);
//       } else cell.t = 's';
//
//       ws[cell_ref] = cell;
//       C++;
//     }
//   }
//   if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
//   return ws;
// }
//
//
// function convertToXLS(data) {
//   var ws = sheetFromArrayOfObjects(data);
//   var ws_name = "SheetJS";
//
//   function Workbook() {
//     if (!(this instanceof Workbook)) return new Workbook();
//     this.SheetNames = [];
//     this.Sheets = {};
//   }
//
//   var wb = new Workbook(),
//     ws = sheetFromArrayOfObjects(data);
//
//   /* add worksheet to workbook */
//   wb.SheetNames.push(ws_name, "Michael");
//   wb.Sheets[ws_name] = ws;
//   var wbout = XLSX.write(wb, {
//     bookType: 'xlsx',
//     bookSST: true,
//     type: 'binary'
//   });
//
//   function stringToArrayBuffer(s) {
//     var buf = new ArrayBuffer(s.length);
//     var view = new Uint8Array(buf);
//     for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
//     return buf;
//   }
//   saveAs(new Blob([stringToArrayBuffer(wbout)], {
//     type: "application/octet-stream"
//   }), "test.xlsx")
// }
// var jsonArr = [{
//   foo: 'bar',
//   qux: 'moo',
//   poo: 123,
//   stux: new Date()
// }, {
//   foo: 'bar',
//   qux: 'moo',
//   poo: 345,
//   stux: new Date()
// }];
// window.onload = function() {
//   console.log("ASD");
//   console.log(XLSX);
//   convertToXLS(jsonArr);
//   console.log("AFTER");
// }









// function doStuffWithDOM(domContent) {
//   console.log("I received the following DOM content:\n" + domContent);
//   alert("SUCCESS FROM BACK");
// }
// chrome.browserAction.onClicked.addListener(function(tab) {
//   chrome.tabs.sendMessage(tab.id, {
//       text: "report_back"
//     },
//     doStuffWithDOM);
// });
