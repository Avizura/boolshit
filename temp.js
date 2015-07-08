// var serverAddress = 'http://192.168.0.168:5000';
// var request = new XMLHttpRequest();

// function toUrlEncoded(obj) {
//   var urlEncoded = "";
//   for (var key in obj) {
//     urlEncoded += encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]) + '&';
//   }
//   return urlEncoded;
// }

console.log("new");

function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}
var ws_name;

chrome.runtime.onMessage.addListener(function(msg, sender, sendResponse) {
  console.log("EURIKA!!");
  console.log(msg);
  if (msg.tab.url == "http://localhost:3000/#/feedback") {
    ws_name = "SheetJS";
  } else if (msg.tab.url == "http://localhost:3000/#/test") {
    ws_name = "Michael";
  }

  function sheetFromArrayOfObjects(data, opts) {
    var ws = {};
    var range = {
      s: {
        c: 10000000,
        r: 10000000
      },
      e: {
        c: 0,
        r: 0
      }
    };
    var C;
    for (var R = 0; R != data.length; ++R) {
      C = 0;
      for (var key in data[R]) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;
        var cell = {
          v: data[R][key]
        };
        if (cell.v == null) {
          cell.v = " ";
        }
        var cell_ref = XLSX.utils.encode_cell({
          c: C,
          r: R
        });

        if (typeof cell.v === 'number') cell.t = 'n';
        else if (typeof cell.v === 'boolean') cell.t = 'b';
        else if (cell.v instanceof Date) {
          cell.t = 'n';
          cell.z = XLSX.SSF._table[14];
          cell.v = datenum(cell.v);
        } else cell.t = 's';

        ws[cell_ref] = cell;
        C++;
      }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
  }


  function convertToXLS(data) {

    function Workbook() {
      if (!(this instanceof Workbook)) return new Workbook();
      this.SheetNames = [];
      this.Sheets = {};
    }

    var wb = new Workbook(),
      ws = sheetFromArrayOfObjects(data);

    /* add worksheet to workbook */
    wb.SheetNames.push("SheetJS", "Michael");
    wb.Sheets[ws_name] = ws;
    wb.Sheets["Michael"] = ws;
    var wbout = XLSX.write(wb, {
      bookType: 'xlsx',
      bookSST: true,
      type: 'binary'
    });

    function stringToArrayBuffer(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }
    saveAs(new Blob([stringToArrayBuffer(wbout)], {
      type: "application/octet-stream"
    }), "test.xlsx")
  }
  var jsonArr = [{
    foo: 'bar',
    qux: 'moo',
    poo: 123,
    stux: new Date()
  }, {
    foo: 'bar',
    qux: 'moo',
    poo: 345,
    stux: new Date()
  }];

  console.log("ASD");
  console.log(XLSX);
  convertToXLS(jsonArr);
  console.log("AFTER");
  sendResponse();


  // console.log(document.getElementById("footer").innerHTML);
  /* If the received message has the expected format... */
  // if (msg.text && (msg.text == "report_back")) {
  //   var error = {
  //     // msg: document.getElementById("footer").innerHTML
  //     msg: "Everything is excellent!"
  //   };
  //   request.onreadystatechange = function() {
  //     if (request.readyState == 4 && request.status == 200) {
  //       console.log(request.responseText);
  //       console.log('wtf mannnnnnnnnnnnn');
  //       sendResponse();
  //       // setTimeout('sendResponse(document.all[0].outerHTML)', 5000);
  //     }
  //   }
  //   request.open('POST', serverAddress + '/mytest', true);
  //   request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  //   request.send(toUrlEncoded(error));
  // }
});





































// var serverAddress = 'http://192.168.0.168:5000';
// var request = new XMLHttpRequest();

// function toUrlEncoded(obj) {
//   var urlEncoded = "";
//   for (var key in obj) {
//     urlEncoded += encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]) + '&';
//   }
//   return urlEncoded;
// }

console.log("new");

function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}
var curName;
var data = [];
var wb;
var wbout;
console.log(curName);

function sheetFromArrayOfObjects(data, opts) {
  var ws = {};
  var range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  };
  var C;
  for (var R = 0; R != data.length; ++R) {
    C = 0;
    for (var key in data[R]) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      var cell = {
        v: data[R][key]
      };
      if (cell.v == null) {
        cell.v = " ";
      }
      var cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R
      });

      if (typeof cell.v === 'number') cell.t = 'n';
      else if (typeof cell.v === 'boolean') cell.t = 'b';
      else if (cell.v instanceof Date) {
        cell.t = 'n';
        cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      } else cell.t = 's';

      ws[cell_ref] = cell;
      C++;
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}


function convertToXLS(data) {

  function Workbook(wb) {
    if (wb instanceof Workbook) return wb;
    this.SheetNames = ["First Page", "Second Page", "Third Page"];
    this.Sheets = {"First Page": "", "Second Page": ""};
  }
  console.log("INSTANCE OF");
  console.log(wb instanceof Workbook);
  console.log(wb);
  wb = new Workbook(wb);
  console.log("AFTER WB");
  console.log(wb);
  var ws = sheetFromArrayOfObjects(data);

  /* add worksheet to workbook */
  wb.Sheets[curName] = ws;
}

var save = function(wbout) {
  function stringToArrayBuffer(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: true,
    type: 'binary'
  });
  saveAs(new Blob([stringToArrayBuffer(wbout)], {
    type: "application/octet-stream"
  }), "test.xlsx");
};

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

// console.log("ASD");
// console.log(XLSX);
// convertToXLS(jsonArr);
// console.log("AFTER");
// sendResponse();



chrome.runtime.onMessage.addListener(function(msg, sender, sendResponse) {
  console.log("EURIKA!!");
  console.log(curName);
  console.log("WB");
  console.log(wb);
  // console.log(msg);
  if (msg.tab.url == "http://localhost:3000/#/feedback") {
    curName = "First Page";
    data = [{
      name: 'Michael',
      surname: 'Dovzhenko',
      age: 20,
      stux: new Date()
    }];
    convertToXLS(data);
    sendResponse();
  } else if (msg.tab.url == "http://localhost:3000/#/test") {
    curName = "Second Page";
    data = [{
      name: 'Vlad',
      surname: 'Dovzhenko',
      age: 23,
      stux: new Date()
    }];
    convertToXLS(data);
    save(wbout);
    sendResponse();
  }
  // console.log(document.getElementById("footer").innerHTML);
  /* If the received message has the expected format... */
  // if (msg.text && (msg.text == "report_back")) {
  //   var error = {
  //     // msg: document.getElementById("footer").innerHTML
  //     msg: "Everything is excellent!"
  //   };
  //   request.onreadystatechange = function() {
  //     if (request.readyState == 4 && request.status == 200) {
  //       console.log(request.responseText);
  //       console.log('wtf mannnnnnnnnnnnn');
  //       sendResponse();
  //       // setTimeout('sendResponse(document.all[0].outerHTML)', 5000);
  //     }
  //   }
  //   request.open('POST', serverAddress + '/mytest', true);
  //   request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  //   request.send(toUrlEncoded(error));
  // }
});
















// var serverAddress = 'http://192.168.0.168:5000';
// var request = new XMLHttpRequest();

// function toUrlEncoded(obj) {
//   var urlEncoded = "";
//   for (var key in obj) {
//     urlEncoded += encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]) + '&';
//   }
//   return urlEncoded;
// }

console.log("new");

function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}
var curName;
var data = [];
var wb = new Workbook();
var wbout;
console.log(curName);

function sheetFromArrayOfObjects(data, opts) {
  var ws = {};
  var range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  };
  var C;
  for (var R = 0; R != data.length; ++R) {
    C = 0;
    for (var key in data[R]) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      var cell = {
        v: data[R][key]
      };
      if (cell.v == null) {
        cell.v = " ";
      }
      var cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R
      });

      if (typeof cell.v === 'number') cell.t = 'n';
      else if (typeof cell.v === 'boolean') cell.t = 'b';
      else if (cell.v instanceof Date) {
        cell.t = 'n';
        cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      } else cell.t = 's';

      ws[cell_ref] = cell;
      C++;
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}

function Workbook() {
  // if (wb instanceof Workbook) return wb;
  this.SheetNames = ["First Page", "Second Page", "Third Page"];
  this.Sheets = {};
}

function convertToXLS(data) {
  // console.log("INSTANCE OF");
  // console.log(wb instanceof Workbook);
  // console.log("AFTER WB");
  // console.log(wb);
  if (!wb)
    wb = new Workbook();
  var ws = sheetFromArrayOfObjects(data);

  /* add worksheet to workbook */
  wb.Sheets[curName] = ws;
  console.log("WB");
  console.log(wb);
}

var save = function(wbout) {
  function stringToArrayBuffer(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: true,
    type: 'binary'
  });
  saveAs(new Blob([stringToArrayBuffer(wbout)], {
    type: "application/octet-stream"
  }), "test.xlsx");
  wb = "";
};

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

// console.log("ASD");
// console.log(XLSX);
// convertToXLS(jsonArr);
// console.log("AFTER");
// sendResponse();



chrome.runtime.onMessage.addListener(function(msg, sender, sendResponse) {
  console.log("EURIKA!!");
  console.log(curName);
  console.log("WB");
  console.log(wb);
  // console.log(msg);
  if (msg.tab.url == "http://localhost:3000/#/feedback") {
    curName = "First Page";
    data = [{
      name: 'Michael',
      surname: 'Dovzhenko',
      age: 20,
      stux: new Date()
    }];
    convertToXLS(data);
    sendResponse();
  } else if (msg.tab.url == "http://localhost:3000/#/test") {
    curName = "Second Page";
    data = [{
      name: 'Vlad',
      surname: 'Dovzhenko',
      age: 23,
      stux: new Date()
    }];
    convertToXLS(data);
    sendResponse();
  } else if (msg.tab.url == "http://localhost:3000/#/history") {
    curName = "Third Page";
    data = [{
      card: "Диагностическая карта"
    }, {
      address: "299053г. Севастополь ул. Вакуленчука 41/9 ИП Пополитов Руслан Анатольевич"
    }, {
      firstTest: "Первичная проверка:",
      firstTestValue: "X",
      retest: "Повторная проверка:",
      retestValue: "",
      mileage: "Пробег ТС:",
      mileageValue: 1000000
    }, {
      regPlate: "Регистр. знак ТС:",
      regPlateValue: "PFDSFDS21",
      carMark: "Марка, модель ТС:",
      carMarkValue: "MITSUBISHI LANCER",
      fuelType: "Тип топлива:",
      fuelTypeValue: "Бензин"
    }, {
      vin: "VIN:",
      vinValue: "FDSFDSFDSASD",
      carCategory: "Категория ТС:",
      carCategoryValue: "M1",
      brakeSystem: "Тип тормозной системы:",
      brakeSystemValue: "гидравлическая"
    }, {
      chassis: "Номер рамы, шасси:",
      chassisValue: "Отсутствует",
      year: "Год выпуска ТС:",
      yearValue: 2006,
      maxWeight: "Разрешенная максимальная масса",
      maxWeightValue: 1750
    }, {
      body: "Номер кузова:",
      bodyValue: "FDSR43fds1",
      tire: "Марка шин:",
      tireValue: "MENTOR",
      weight: "Масса без нагрузки:",
      weightValue: 1232
    }, {
      wtf: "СРТС или ПТС (серия, номер, выдан, кем, когда):",
      wtfValue: "9225 # ПОАЫОЫВАВЛЫОАРЫВЛОАРЫВЛОАОЫВАРВЫ"
    }, {
      owner: "Владелец ТС:",
      ownerValue: ""
    }, {
      humidity: "Влажность, %",
      pressure: "Давление, кПА",
      temperature: "Температура, С",
      windSpeed: "Скорость ветра, м/с"
    }, {
      humidityValue: "",
      pressureValue: "",
      temperatureValue: "",
      windSpeedValue: ""
    }, {
      number: "№ в ДК",
      value: "Параметры и требования, предъявляемые к транспортным средствам при проведении технического контроля",
      yeap: "Наличие соответствия"
    }, {
      numberValue: "65.",
      value: "Установка государственных регистрационных знаков в соответствии с требованиями",
      yeapValue: ""
    }, {
      numberValue: "41.",
      value: "Работоспособность замков дверей кузова,  кабины, механизмов регулировки и фиксирующих устройств сидений, устройства обогрева и обдува ветрового стекла, противоугонного устройства",
      yeapValue: ""
    }, {
      numberValue: "55.",
      value: "Наличие знака аварийной остановки",
      yeapValue: ""
    }, {
      numberValue: "57.",
      value: "Наличие огнетушителей, соответствующих установленным требованиям",
      yeapValue: ""
    }, {
      numberValue: "56.",
      value: "Наличие не менее двух противооткатных упоров",
      yeapValue: ""
    }, {
      numberValue: "58.",
      value: "Надежное крепление поручней в автобусах, запасного колеса, аккумуляторной батареи, сидений, огнетушителей и медицинской аптечки",
      yeapValue: ""
    }, {
      numberValue: "60.",
      value: "Наличие над колёсных грязезащитных устройств, отвечающих установленным требованиям",
      yeapValue: ""
    }, {
      numberValue: "62.",
      value: "Работоспособность держателя запасного колеса, лебедки и механизма подъема-опускания запасного колеса",
      yeapValue: ""
    }, {
      numberValue: "46.",
      value: "Наличие обозначений аварийных выходов и табличек по правилам их использования. Обеспечение свободного доступа к аварийным выходам",
      yeapValue: ""
    }, {
      numberValue: "43.",
      value: "Работоспособность аварийного выключателя дверей и сигнала требования остановки",
      yeapValue: ""
    }, {
      numberValue: "44.",
      value: "Работоспособность аварийных выходов, приборов внутреннего освещения салона, привода управления дверями и сигнализации их работы",
      yeapValue: ""
    }, {
      numberValue: "59.",
      value: "Работоспособность механизмов регулировки сидений",
      yeapValue: ""
    }, {
      numberValue: "54.",
      value: "Оснащение транспортных средств исправными ремнями безопасности",
      yeapValue: ""
    }, {
      numberValue: "45.",
      value: "Наличие работоспособного звукового сигнального прибора",
      yeapValue: ""
    }, {
      numberValue: "47.",
      value: "Наличие задних и боковых  защитных устройств, соответствие их нормам",
      yeapValue: ""
    }, {
      numberValue: "42.",
      value: "Работоспособность запоров бортов грузовой платформы и запоров горловин цистерн",
      yeapValue: ""
    }, {
      numberValue: "48.",
      value: "Работоспособность автоматического замка, ручной и автоматической блокировки седельно-сцепного устройства. Отсутствие видимых повреждений сцепных устройств",
      yeapValue: ""
    }, {
      numberValue: "49.",
      value: "Наличие работоспособных предохранительных приспособлений у одноосных прицепов (за исключением роспусков) и прицепов, не оборудованных рабочей тормозной системой",
      yeapValue: ""
    }, {
      numberValue: "50.",
      value: "Оборудование прицепов (за исключением одноосных и роспусков) исправным устройством, поддерживающим сцепную петлю дышла в положении, облегчающем сцепку и расцепку с тяговым автомобилем",
      yeapValue: ""
    }, {
      numberValue: "51.",
      value: "Отсутствие продольного люфта в беззазорных тягово-сцепных устройствах с тяговой вилкой для сцепленного с прицепом тягача",
      yeapValue: ""
    }, {
      numberValue: "52.",
      value: "Обеспечение тягово-сцепными устройствами легковых автомобилей беззазорной сцепки сухарей замкового устройства с шаром",
      yeapValue: ""
    }, {
      numberValue: "53.",
      value: "Соответствие размерных характеристик сцепных устройств установленным требованиям",
      yeapValue: ""
    }, {
      numberValue: "61.",
      value: "Соответствие вертикальной статической нагрузки на тяговое устройство автомобиля от сцепной петли одноосного прицепа (прицепа-роспуска) нормам",
      yeapValue: ""
    }, {
      numberValue: "63.",
      value: "Работоспособность механизмов подъема и опускания опор и фиксаторов транспортного положения опор",
      yeapValue: ""
    }, {
      numberValue: "64.",
      value: "Соответствие каплепадения масел и рабочих жидкостей нормам",
      yeapValue: ""
    }, {
      numberValue: "33.",
      value: "Отсутствие подтекания и каплепадения топлива в системе питания",
      yeapValue: ""
    }, {
      numberValue: "34.",
      value: "Работоспособность запорных устройств и устройств перекрытия топлива",
      yeapValue: ""
    }, {
      numberValue: "37.",
      value: "Наличие зеркал заднего вида в соответствии с требованиями",
      yeapValue: ""
    }, {
      numberValue: "40.",
      value: "Отсутствие трещин на ветровом стекле в зоне очистки водительского стеклоочистителя",
      yeapValue: ""
    }, {
      numberValue: "38.",
      value: "Отсутствие дополнительных предметов или покрытий, ограничивающих обзорность с места водителя. Соответствие полосы пленки в  верхней  части ветрового  стекла  установленным требованиям",
      yeapValue: ""
    }, {
      wipers: "Стеклоочистители и стеклоомыватели"
    }, {
      numberValue: "23.",
      value: "Наличие  стеклоочистителя и форсунки стеклоомывателя ветрового стекла",
      yeapValue: ""
    }, {
      numberValue: "24.",
      value: "Обеспечение стеклоомывателем подачи жидкости в зоны очистки стекла",
      yeapValue: ""
    }, {
      numberValue: "25.",
      value: "Работоспособность стеклоочистителей составляет не менее 35 двойных ходов щеток/мин и стеклоомывателей",
      yeapValue: ""
    }, {
      wheels: "Шины и колеса"
    }, {
      numberValue: "26.",
      value: "Соответствие высоты рисунка протектора шин установленным требованиям     1,6 мм для М1 и прицепов к ним, 1,0 мм для N и прицепов к ним и 2,0 мм для М2, М3",
      yeapValue: ""
    }, {
      numberValue: "27.",
      value: "Отсутствие признаков непригодности шин к эксплуатации",
      yeapValue: ""
    }, {
      numberValue: "28.",
      value: "Наличие всех болтов или гаек крепления дисков и ободьев колес",
      yeapValue: ""
    }, {
      numberValue: "29.",
      value: "Отсутствие трещин на дисках и ободьях колес",
      yeapValue: ""
    }, {
      numberValue: "30.",
      value: "Отсутствие видимых нарушений формы и размеров крепежных отверстий в дисках колес",
      yeapValue: ""
    }, {
      numberValue: "31.",
      value: "Установка шин на транспортное средство в соответствии с требованиями",
      yeapValue: ""
    }];
    convertToXLS(data);
    save(wbout);
    sendResponse();
  }
  // console.log(document.getElementById("footer").innerHTML);
  /* If the received message has the expected format... */
  // if (msg.text && (msg.text == "report_back")) {
  //   var error = {
  //     // msg: document.getElementById("footer").innerHTML
  //     msg: "Everything is excellent!"
  //   };
  //   request.onreadystatechange = function() {
  //     if (request.readyState == 4 && request.status == 200) {
  //       console.log(request.responseText);
  //       console.log('wtf mannnnnnnnnnnnn');
  //       sendResponse();
  //       // setTimeout('sendResponse(document.all[0].outerHTML)', 5000);
  //     }
  //   }
  //   request.open('POST', serverAddress + '/mytest', true);
  //   request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  //   request.send(toUrlEncoded(error));
  // }
});


// data = [{
//   "Диагностическая карта": ""
// }, {
//   "299053г. Севастополь ул. Вакуленчука 41/9 ИП Пополитов Руслан Анатольевич": ""
// }, {
//   "Первичная проверка:": "X", "Повторная проверка:": "", "Пробег ТС:" : 1000000
// }, {
//   "Регистр. знак ТС:": "PFDSFDS21", "Марка, модель ТС:": "MITSUBISHI LANCER", "Тип топлива:": "Бензин",
//   "VIN:": "FDSFDSFDSASD", "Категория ТС:" : "M1", "Тип тормозной системы:": "гидравлическая"
// }]










































var serverAddress = 'http://127.0.0.1:5000';
var request = new XMLHttpRequest();

function toUrlEncoded(obj) {
  var urlEncoded = "";
  for (var key in obj) {
    urlEncoded += encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]) + '&';
  }
  return urlEncoded;
}

function ge(id) {
  return document.getElementById(id).innerHTML;
}

function geName(name) {
  return document.getElementsByName(name)[0].innerHTML;
}

function getSelectedValue(id) {
  var element = ge(id);
  return element.options[element.selectedIndex].value;
}

function getStep2Value(id) {
  var el = getElementById('parameter_row_' + id);
  if (el.classList.contains('tr_sel_disabled') || el.classList.contains('nonclickable')) {
    return '+';
  } else if (el.classList.contains('tr_sel_red')) return '';
}

console.log("new");

chrome.runtime.onMessage.addListener(function(msg, sender, sendResponse) {
  console.log("EURIKA!!");
  console.log(sender);
  var data;
  var step = document.getElementsByClassName("sel");
  if (step[2].innerHTML == 'Исходные данные') {
    var msg = {
      msg: "step 1",
      firstName: ge("I_LICA"),
      secondName: ge("F_LICA"),
      lastName: ge("O_LICA"),
      vin: ge('VIN'),
      year: ge('GOD'),
      mark: ge('MARKA'),
      model: ge('MODEL'),
      body: ge('NOMER_KUZOVA'),
      chassis: ge('NOMER_RAMY_ID'),
      mileage: geName('PROBEG'),
      regPlate: ge('REG_ZNAK'),
      weight: geName('MASSA_BEZ_NAGRUZKI'),
      vehicleCategory: getSelectedValue('KATEG_ID'),
      tire: getSelectedValue('MARKA_SHIN'),
      maxWeight: geName('RAZRESH_MAKS_MASSA'),
      fuel: getSelectedValue('TIP_TOPLIVA'),
      brakeSystem: getSelectedValue('TIP_TORMOZ_SISTEMY'),
      SVID_O_REG_SERIA: ge('SVID_O_REG_SERIA'),
      SVID_O_REG_NOMER: ge('SVID_O_REG_NOMER'),
      SVID_O_REG_KOGDA: ge('SVID_O_REG_KOGDA'),
      SVID_O_REG_KEM: geName('SVID_O_REG_KEM')
    };
  } else if (step[0].innerHTML == 'Результат диагностики') {
    data = {
      _1: getStep2Value(1),
      _2: getStep2Value(2),
      _3: getStep2Value(3),
      _4: getStep2Value(4),
      _5: getStep2Value(5),
      _6: getStep2Value(6),
      _7: getStep2Value(7),
      _8: getStep2Value(8),
      _9: getStep2Value(9),
      _10: getStep2Value(10),
      _11: getStep2Value(11),
      _12: getStep2Value(12),
      _13: getStep2Value(13),
      _14: getStep2Value(14),
      _15: getStep2Value(15),
      _16: getStep2Value(16),
      _17: getStep2Value(17),
      _18: getStep2Value(18),
      _19: getStep2Value(19),
      _20: getStep2Value(20),
      _21: getStep2Value(21),
      _22: getStep2Value(22),
      _23: getStep2Value(23),
      _24: getStep2Value(24),
      _25: getStep2Value(25),
      _26: getStep2Value(26),
      _27: getStep2Value(27),
      _28: getStep2Value(28),
      _29: getStep2Value(29),
      _30: getStep2Value(30),
      _31: getStep2Value(31),
      _32: getStep2Value(32),
      _33: getStep2Value(33),
      _34: getStep2Value(34),
      _35: getStep2Value(35),
      _36: getStep2Value(36),
      _37: getStep2Value(37),
      _38: getStep2Value(38),
      _39: getStep2Value(39),
      _40: getStep2Value(40),
      _41: getStep2Value(41),
      _42: getStep2Value(42),
      _43: getStep2Value(43),
      _44: getStep2Value(44),
      _45: getStep2Value(45),
      _46: getStep2Value(46),
      _47: getStep2Value(47),
      _48: getStep2Value(48),
      _49: getStep2Value(49),
      _50: getStep2Value(50),
      _51: getStep2Value(51),
      _52: getStep2Value(52),
      _53: getStep2Value(53),
      _54: getStep2Value(54),
      _55: getStep2Value(55),
      _56: getStep2Value(56),
      _57: getStep2Value(57),
      _58: getStep2Value(58),
      _59: getStep2Value(59),
      _60: getStep2Value(60),
      _61: getStep2Value(61),
      _62: getStep2Value(62),
      _63: getStep2Value(63),
      _64: getStep2Value(64),
      _64: getStep2Value(65)
    }
  } else if (step[2].innerHTML == 'Заключение') {
    data = {
      msg: 'step 3',
      SROK_DEISTV: ge('SROK_DEISTV')
    };
  }
  //if(активный шаг) {}
  // var msg = {
  //   msg: "Everything is excellent!",
  //   firstName: ge("I_LICA"),
  //   secondName: ge("F_LICA"),
  //   lastName: ge("O_LICA"),
  //   vin: ge('VIN'),
  //   year: ge('GOD'),
  //   mark: ge('MARKA'),
  //   model: ge('MODEL'),
  //   body: ge('NOMER_KUZOVA'),
  //   chassis: ge('NOMER_RAMY_ID'),
  //   mileage: geName('PROBEG'),
  //   regPlate: ge('REG_ZNAK'),
  //   weight: geName('MASSA_BEZ_NAGRUZKI'),
  //   vehicleCategory: getSelectedValue('KATEG_ID'),
  //   tire: getSelectedValue('MARKA_SHIN'),
  //   maxWeight: geName('RAZRESH_MAKS_MASSA'),
  //   fuel: getSelectedValue('TIP_TOPLIVA'),
  //   brakeSystem: getSelectedValue('TIP_TORMOZ_SISTEMY'),
  //   SVID_O_REG_SERIA: ge('SVID_O_REG_SERIA'),
  //   SVID_O_REG_NOMER: ge('SVID_O_REG_NOMER'),
  //   SVID_O_REG_KOGDA: ge('SVID_O_REG_KOGDA'),
  //   SVID_O_REG_KEM: geName('SVID_O_REG_KEM')
  // };
  // if (msg.tab.url == "http://localhost:3000/#/feedback") {
  //   data = {
  //     msg: 'step 1',
  //     firstName: 'Michael',
  //     secondName: 'Dovzhenko',
  //     lastName: 'Igorevich',
  //     vin: '41421321',
  //     year: '1995',
  //     mark: 'HYINADAI',
  //     model: 'SONATA',
  //     body: 'XZ',
  //     chassis: 'WTF',
  //     mileage: '1500',
  //     regPlate: '1243FSD43',
  //     weight: '4212',
  //     vehicleCategory: 'CAR',
  //     tire: 'BMW',
  //     maxWeight: '41242',
  //     fuel: 'BENZIN',
  //     brakeSystem: 'ARTEEZY',
  //     regSERIA: '9225',
  //     regNOMER: '63543',
  //     regKOGDA: '15.09.2014',
  //     regKEM: 'МРЭО ГИБДД УМВД РОССИИ ПО Г. СЕВАСТОПОЛЮ'
  //   };
  // } else if (msg.tab.url == "http://localhost:3000/#/test") {
    // data = {
    //   _1: getStep2Value(1),
    //   _2: getStep2Value(2),
    //   _3: getStep2Value(3),
    //   _4: getStep2Value(4),
    //   _5: getStep2Value(5),
    //   _6: getStep2Value(6),
    //   _7: getStep2Value(7),
    //   _8: getStep2Value(8),
    //   _9: getStep2Value(9),
    //   _10: getStep2Value(10),
    //   _11: getStep2Value(11),
    //   _12: getStep2Value(12),
    //   _13: getStep2Value(13),
    //   _14: getStep2Value(14),
    //   _15: getStep2Value(15),
    //   _16: getStep2Value(16),
    //   _17: getStep2Value(17),
    //   _18: getStep2Value(18),
    //   _19: getStep2Value(19),
    //   _20: getStep2Value(20),
    //   _21: getStep2Value(21),
    //   _22: getStep2Value(22),
    //   _23: getStep2Value(23),
    //   _24: getStep2Value(24),
    //   _25: getStep2Value(25),
    //   _26: getStep2Value(26),
    //   _27: getStep2Value(27),
    //   _28: getStep2Value(28),
    //   _29: getStep2Value(29),
    //   _30: getStep2Value(30),
    //   _31: getStep2Value(31),
    //   _32: getStep2Value(32),
    //   _33: getStep2Value(33),
    //   _34: getStep2Value(34),
    //   _35: getStep2Value(35),
    //   _36: getStep2Value(36),
    //   _37: getStep2Value(37),
    //   _38: getStep2Value(38),
    //   _39: getStep2Value(39),
    //   _40: getStep2Value(40),
    //   _41: getStep2Value(41),
    //   _42: getStep2Value(42),
    //   _43: getStep2Value(43),
    //   _44: getStep2Value(44),
    //   _45: getStep2Value(45),
    //   _46: getStep2Value(46),
    //   _47: getStep2Value(47),
    //   _48: getStep2Value(48),
    //   _49: getStep2Value(49),
    //   _50: getStep2Value(50),
    //   _51: getStep2Value(51),
    //   _52: getStep2Value(52),
    //   _53: getStep2Value(53),
    //   _54: getStep2Value(54),
    //   _55: getStep2Value(55),
    //   _56: getStep2Value(56),
    //   _57: getStep2Value(57),
    //   _58: getStep2Value(58),
    //   _59: getStep2Value(59),
    //   _60: getStep2Value(60),
    //   _61: getStep2Value(61),
    //   _62: getStep2Value(62),
    //   _63: getStep2Value(63),
    //   _64: getStep2Value(64),
    //   _64: getStep2Value(65)
    // }
    // data = {
    //   msg: 'step 2',
    //   _1: '+',
    //   _2: '+',
    //   _3: '+',
    //   _4: '+',
    //   _5: '+',
    //   _6: '+',
    //   _7: '+',
    //   _8: '+',
    //   _9: '+',
    //   _10: '+',
    //   _11: '+',
    //   _12: '+',
    //   _13: '+',
    //   _14: '+',
    //   _15: '+',
    //   _16: '+',
    //   _17: '+',
    //   _18: '+',
    //   _19: '+',
    //   _20: '+',
    //   _21: '+',
    //   _22: '+',
    //   _23: '+',
    //   _24: '+',
    //   _25: '+',
    //   _26: '+',
    //   _27: '+',
    //   _28: '+',
    //   _29: '+',
    //   _30: '+',
    //   _31: '+',
    //   _32: '+',
    //   _33: '+',
    //   _34: '+',
    //   _35: '+',
    //   _36: '+',
    //   _37: '+',
    //   _38: '+',
    //   _39: '+',
    //   _40: '+',
    //   _41: '+',
    //   _42: '+',
    //   _43: '+',
    //   _44: '+',
    //   _45: '+',
    //   _46: '+',
    //   _47: '+',
    //   _48: '+',
    //   _49: '+',
    //   _50: '+',
    //   _51: '+',
    //   _52: '+',
    //   _53: '+',
    //   _54: '+',
    //   _55: '+',
    //   _56: '+',
    //   _57: '+',
    //   _58: '+',
    //   _59: '+',
    //   _60: '+',
    //   _61: '+',
    //   _62: '+',
    //   _63: '+',
    //   _64: '+',
    //   _64: '+',
    //   _65: '+'
    // }
  // } else if (msg.tab.url == "http://localhost:3000/#/history") {
    // data = {
    //   msg: 'step 3',
    //   SROK_DEISTV: ge('SROK_DEISTV'),
    //
    // };
  // }
  request.onreadystatechange = function() {
    if (request.readyState == 4 && request.status == 200) {
      console.log(request.responseText);
      console.log('response received');
      sendResponse({
        farewell: "goodbye"
      });
    }
  }
  request.open('POST', serverAddress + '/mytest', true);
  request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  request.send(toUrlEncoded(data));
});
