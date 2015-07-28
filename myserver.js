var http = require('http');
var url = require('url');
var qs = require('querystring');
var fs = require('fs');
var path = require('path');
var serverPort = '5000';
var serverAddress = '127.0.0.1';
var journal, wb, wsOpts, ws0, ws, ws2, ws3, ws4, data1, diagnosticIssues, failedRequirements, notes, regNumber;
try {
  var xl = require('excel4node');
} catch (e) {
  var xl = require('./node_modules/excel4node/lib/index.js');
}
var wbPath = '' + fs.readFileSync('path.txt');
var journalPath = '' + fs.readFileSync('journalPath.txt');

function createWorkBook(post) {
  var options = {
    jszip: {
      compression: 'DEFLATE'
    }
  }
  wb = new xl.WorkBook(options);
  journal = new xl.WorkBook(options);

  wsOpts = {
    margins: {
      left: 15
    },
    printOptions: {
      centerHorizontal: true,
      centerVertical: true
    },
    view: {
      zoom: 100
    },
    fitToPage: {
      fitToHeight: 100,
      fitToWidth: 100,
      orientation: 'landscape'
    }
  }

  ws0 = journal.WorkSheet('Журнал регистрации ТС', wsOpts);
  ws = wb.WorkSheet('Первичный', wsOpts);
  /*
    Styles
  */
  var myStyle = wb.Style();
  myStyle.Font.Family('Times New Roman');
  myStyle.Font.Alignment.Vertical('center');
  myStyle.Font.Alignment.Horizontal('left');
  myStyle.Font.WrapText(true);

  var myStyle2 = wb.Style();
  myStyle2.Font.Family('Times New Roman');
  myStyle2.Font.Alignment.Vertical('center');
  myStyle2.Font.Alignment.Horizontal('center');
  myStyle2.Fill.Color('FFEAEA');
  myStyle2.Fill.Pattern('solid');
  myStyle2.Font.WrapText(true);
  myStyle2.Border({
    left: {
      style: 'thin',
      color: 'D0D0D0'
    },
    right: {
      style: 'thin',
      color: 'D0D0D0'
    },
    top: {
      style: 'thin',
      color: 'D0D0D0'
    },
    bottom: {
      style: 'thin',
      color: 'D0D0D0'
    }
  });

  var leftBorder = wb.Style();
  leftBorder.Border({
    left: {
      style: 'thick',
      color: '0000FF'
    }
  });
  var bottomBorder = wb.Style();
  bottomBorder.Border({
    bottom: {
      style: 'thick',
      color: '0000FF'
    }
  });

  var myStyle4 = wb.Style();
  myStyle4.Font.Alignment.Vertical('center');
  myStyle4.Font.Alignment.Horizontal('left');
  myStyle4.Font.Family('Times New Roman');
  myStyle4.Font.Size(10);
  myStyle4.Font.WrapText();

  var myStyle4Bold = wb.Style();
  myStyle4Bold.Font.Alignment.Vertical('center');
  myStyle4Bold.Font.Alignment.Horizontal('left');
  myStyle4Bold.Font.Family('Times New Roman');
  myStyle4Bold.Font.Size(10);
  myStyle4Bold.Font.Bold();
  myStyle4Bold.Font.WrapText();


  var myStyle5 = wb.Style();
  myStyle5.Font.Alignment.Vertical('center');
  myStyle5.Font.Alignment.Horizontal('center');
  myStyle5.Font.Family('Times New Roman');
  myStyle5.Font.Size(10);
  myStyle5.Font.WrapText();

  var myStyle6 = wb.Style();
  myStyle6.Font.Family('Times New Roman');
  myStyle6.Font.Alignment.Vertical('center');
  myStyle6.Font.Alignment.Horizontal('left');
  myStyle6.Fill.Color('FFEAEA');
  myStyle6.Fill.Pattern('solid');
  myStyle6.Border({
    left: {
      style: 'thin',
      color: 'D0D0D0'
    },
    right: {
      style: 'thin',
      color: 'D0D0D0'
    },
    top: {
      style: 'thin',
      color: 'D0D0D0'
    },
    bottom: {
      style: 'thin',
      color: 'D0D0D0'
    }
  });

  /*
    Code to generate page Первичный
  */
  //ширина столбцов
  ws.Column(1).Width(5);
  ws.Column(2).Width(15);
  ws.Column(3).Width(10);
  ws.Column(4).Width(20);
  ws.Column(5).Width(20);
  ws.Column(6).Width(10);
  ws.Column(7).Width(15);
  ws.Column(8).Width(12);
  ws.Column(9).Width(13);
  ws.Column(10).Width(8);
  ws.Column(11).Width(8);
  ws.Column(12).Width(12);

  //высота строк
  ws.Row(14).Height(30);
  ws.Row(31).Height(30);
  ws.Row(42).Height(30);
  ws.Row(55).Height(30);
  ws.Row(70).Height(30);

  //объединение ячеек
  for (var i = 3; i < 8; i++) {
    ws.Cell(i, 1, i, 2, true);
    ws.Cell(i, 3, i, 4, true);
    ws.Cell(i, 6, i, 7, true);
    ws.Cell(i, 8, i, 9, true);
    ws.Cell(i, 10, i, 12, true);
  }
  ws.Cell(8, 1, 8, 4, true);
  ws.Cell(8, 5, 8, 12, true);
  ws.Cell(9, 1, 9, 2, true);
  ws.Cell(9, 3, 9, 12, true);
  for (var i = 10; i < 12; ++i) {
    ws.Cell(i, 2, i, 3, true);
    ws.Cell(i, 4, i, 5, true);
    ws.Cell(i, 6, i, 8, true);
    ws.Cell(i, 9, i, 11, true);
  }
  for (var i = 12; i < 56; ++i)
    ws.Cell(i, 2, i, 11, true);
  ws.Cell(56, 2, 56, 8, true);
  for (var i = 57; i < 79; ++i)
    ws.Cell(i, 2, i, 11, true);
  ws.Cell(79, 1, 79, 2, true);
  ws.Cell(80, 1, 82, 2, true);
  for (var i = 80; i < 82; ++i) {
    for (var j = 3; j < 6; j += 2)
      ws.Cell(i, j, i, j + 1, true);
  }
  ws.Cell(81, 10, 84, 10, true);
  ws.Cell(80, 11, 80, 12, true);
  ws.Cell(81, 11, 84, 12, true);
  for (var i = 83; i < 87; ++i) {
    for (var j = 1; j < 7; j += 2)
      ws.Cell(i, j, i, j + 1, true);
  }
  for (var i = 87; i < 93; ++i)
    ws.Cell(i, 2, i, 11, true);

  ws.Cell(93, 8, 93, 9, true);
  ws.Cell(93, 10, 93, 11, true);
  ws.Cell(94, 1, 94, 3, true);
  ws.Cell(94, 4, 94, 5, true);
  ws.Cell(94, 6, 94, 9, true);
  ws.Cell(94, 10, 94, 12, true);
  ws.Cell(95, 1, 95, 3, true);
  ws.Cell(95, 10, 95, 12, true);

  data = [{}, {
    card: "Диагностическая карта"
  }, {
    address: "299053г. Севастополь ул. Вакуленчука 41/9 ИП Пополитов Руслан Анатольевич"
  }, {
    v1: "Первичная проверка:",
    v2: "",
    v3: "Повторная проверка:",
    v4: "",
    v5: "Пробег ТС:",
    v6: post.mileage
  }, {
    v1: "Регистр. знак ТС:",
    v2: post.regPlate,
    v3: "Марка, модель ТС:",
    v4: post.mark + " " + post.model,
    v5: "Тип топлива:",
    v6: post.fuel
  }, {
    v1: "VIN:",
    v2: post.vin,
    v3: "Категория ТС:",
    v4: post.vehicleCategory,
    v5: "Тип тормозной системы:",
    v6: post.brakeSystem
  }, {
    v1: "Номер рамы, шасси:",
    v2: post.chassis,
    v3: "Год выпуска ТС:",
    v4: post.year,
    v5: "Разрешенная max масса:",
    v6: post.maxWeight
  }, {
    v1: "Номер кузова:",
    v2: post.body,
    v3: "Марка шин:",
    v4: post.tire,
    v5: "Масса без нагрузки:",
    v6: post.weight
  }, {
    v1: "СРТС или ПТС (серия, номер, выдан, кем, когда):",
    v2: post.regSERIA + " №" + post.regNOMER + " " + post.regKEM + " от " + post.regKOGDA
  }, {
    v1: "Владелец ТС:",
    v2: post.secondName + " " + post.firstName + " " + post.lastName
  }, {
    v1: "Влажность, %",
    v2: "Давление, кПА",
    v3: "Температура, С",
    v4: "Скорость ветра, м/с"
  }, {
    v1: "50",
    v2: "40",
    v3: "30",
    v4: "20"
  }, {
    v1: "№ в ДК",
    v2: "Параметры и требования, предъявляемые к транспортным средствам при проведении технического контроля",
    v3: "Наличие соответствия"
  }, {
    v1: "65.",
    v2: "Установка государственных регистрационных знаков в соответствии с требованиями",
    v3: ""
  }, {
    v1: "41.",
    v2: "Работоспособность замков дверей кузова,  кабины, механизмов регулировки и фиксирующих устройств сидений, устройства обогрева и обдува ветрового стекла, противоугонного устройства",
    v3: ""
  }, {
    v1: "55.",
    v2: "Наличие знака аварийной остановки",
    v3: ""
  }, {
    v1: "57.",
    v2: "Наличие огнетушителей, соответствующих установленным требованиям",
    v3: ""
  }, {
    v1: "56.",
    v2: "Наличие не менее двух противооткатных упоров",
    v3: ""
  }, {
    v1: "58.",
    v2: "Надежное крепление поручней в автобусах, запасного колеса, аккумуляторной батареи, сидений, огнетушителей и медицинской аптечки",
    v3: ""
  }, {
    v1: "60.",
    v2: "Наличие над колёсных грязезащитных устройств, отвечающих установленным требованиям",
    v3: ""
  }, {
    v1: "62.",
    v2: "Работоспособность держателя запасного колеса, лебедки и механизма подъема-опускания запасного колеса",
    v3: ""
  }, {
    v1: "46.",
    v2: "Наличие обозначений аварийных выходов и табличек по правилам их использования. Обеспечение свободного доступа к аварийным выходам",
    v3: ""
  }, {
    v1: "43.",
    v2: "Работоспособность аварийного выключателя дверей и сигнала требования остановки",
    v3: ""
  }, {
    v1: "44.",
    v2: "Работоспособность аварийных выходов, приборов внутреннего освещения салона, привода управления дверями и сигнализации их работы",
    v3: ""
  }, {
    v1: "59.",
    v2: "Работоспособность механизмов регулировки сидений",
    v3: ""
  }, {
    v1: "54.",
    v2: "Оснащение транспортных средств исправными ремнями безопасности",
    v3: ""
  }, {
    v1: "45.",
    v2: "Наличие работоспособного звукового сигнального прибора",
    v3: ""
  }, {
    v1: "47.",
    v2: "Наличие задних и боковых  защитных устройств, соответствие их нормам",
    v3: ""
  }, {
    v1: "42.",
    v2: "Работоспособность запоров бортов грузовой платформы и запоров горловин цистерн",
    v3: ""
  }, {
    v1: "48.",
    v2: "Работоспособность автоматического замка, ручной и автоматической блокировки седельно-сцепного устройства. Отсутствие видимых повреждений сцепных устройств",
    v3: ""
  }, {
    v1: "49.",
    v2: "Наличие работоспособных предохранительных приспособлений у одноосных прицепов (за исключением роспусков) и прицепов, не оборудованных рабочей тормозной системой",
    v3: ""
  }, {
    v1: "50.",
    v2: "Оборудование прицепов (за исключением одноосных и роспусков) исправным устройством, поддерживающим сцепную петлю дышла в положении, облегчающем сцепку и расцепку с тяговым автомобилем",
    v3: ""
  }, {
    v1: "51.",
    v2: "Отсутствие продольного люфта в беззазорных тягово-сцепных устройствах с тяговой вилкой для сцепленного с прицепом тягача",
    v3: ""
  }, {
    v1: "52.",
    v2: "Обеспечение тягово-сцепными устройствами легковых автомобилей беззазорной сцепки сухарей замкового устройства с шаром",
    v3: ""
  }, {
    v1: "53.",
    v2: "Соответствие размерных характеристик сцепных устройств установленным требованиям",
    v3: ""
  }, {
    v1: "61.",
    v2: "Соответствие вертикальной статической нагрузки на тяговое устройство автомобиля от сцепной петли одноосного прицепа (прицепа-роспуска) нормам",
    v3: ""
  }, {
    v1: "63.",
    v2: "Работоспособность механизмов подъема и опускания опор и фиксаторов транспортного положения опор",
    v3: ""
  }, {
    v1: "64.",
    v2: "Соответствие каплепадения масел и рабочих жидкостей нормам",
    v3: ""
  }, {
    v1: "33.",
    v2: "Отсутствие подтекания и каплепадения топлива в системе питания",
    v3: ""
  }, {
    v1: "34.",
    v2: "Работоспособность запорных устройств и устройств перекрытия топлива",
    v3: ""
  }, {
    v1: "37.",
    v2: "Наличие зеркал заднего вида в соответствии с требованиями",
    v3: ""
  }, {
    v1: "40.",
    v2: "Отсутствие трещин на ветровом стекле в зоне очистки водительского стеклоочистителя",
    v3: ""
  }, {
    v1: "38.",
    v2: "Отсутствие дополнительных предметов или покрытий, ограничивающих обзорность с места водителя. Соответствие полосы пленки в  верхней  части ветрового  стекла  установленным требованиям",
    v3: ""
  }, {
    v1: "",
    v2: "Стеклоочистители и стеклоомыватели",
    v3: ""
  }, {
    v1: "23.",
    v2: "Наличие  стеклоочистителя и форсунки стеклоомывателя ветрового стекла",
    v3: ""
  }, {
    v1: "24.",
    v2: "Обеспечение стеклоомывателем подачи жидкости в зоны очистки стекла",
    v3: ""
  }, {
    v1: "25.",
    v2: "Работоспособность стеклоочистителей составляет не менее 35 двойных ходов щеток/мин и стеклоомывателей",
    v3: ""
  }, {
    v1: "",
    v2: "Шины и колеса",
    v3: ""
  }, {
    v1: "26.",
    v2: "Соответствие высоты рисунка протектора шин установленным требованиям     1,6 мм для М1 и прицепов к ним, 1,0 мм для N и прицепов к ним и 2,0 мм для М2, М3",
    v3: ""
  }, {
    v1: "27.",
    v2: "Отсутствие признаков непригодности шин к эксплуатации",
    v3: ""
  }, {
    v1: "28.",
    v2: "Наличие всех болтов или гаек крепления дисков и ободьев колес",
    v3: ""
  }, {
    v1: "29.",
    v2: "Отсутствие трещин на дисках и ободьях колес",
    v3: ""
  }, {
    v1: "30.",
    v2: "Отсутствие видимых нарушений формы и размеров крепежных отверстий в дисках колес",
    v3: ""
  }, {
    v1: "31.",
    v2: "Установка шин на транспортное средство в соответствии с требованиями",
    v3: ""
  }, {
    v1: "",
    v2: "Тормозные системы",
    v3: ""
  }, {
    v1: "1.",
    v2: "Соответствие показателей эффективности торможения РТС и устойчивости торможения 0,53 для  М1; 0,43 для М2,М3,N1,N2,N3 ; 0,45 - для пр-ов с 2я и более осями О1-О4; 0,41 - пр-пы с центр. осью и п/п категории О1-О4.",
    v3: ""
  }, {
    v1: "2.",
    v2: "Соответствие разности тормозных сил установленным требованиям( дисковые не более 20% барабанные  не более 25%)",
    v3: "1-я",
    v4: "2-я",
    v5: "3-я",
    v6: "4-я"
  }, {
    v1: "3.",
    v2: "Работоспособность рабочей тормозной системы автопоездов с пневматическим тормозным приводом в режиме аварийного (автоматического) торможения",
    v3: ""
  }, {
    v1: "1.",
    v2: "Соответствие показателей эффективности торможения СТС и устойчивости торможения не менее 0,16",
    v3: ""
  }, {
    v1: "4.",
    v2: "Отсутствие утечек сжатого воздуха из колесных тормозных камер",
    v3: ""
  }, {
    v1: "5.",
    v2: "Отсутствие подтеканий тормозной жидкости, нарушения герметичности трубопроводов или соединений в гидравлическом тормозном приводе",
    v3: ""
  }, {
    v1: "6.",
    v2: "Отсутствие коррозии, грозящей потерей герметичности или разрушением",
    v3: ""
  }, {
    v1: "7.",
    v2: "Отсутствие механических повреждений тормозных трубопроводов",
    v3: ""
  }, {
    v1: "8.",
    v2: "Отсутствие трещин остаточной деформации деталей тормозного привода",
    v3: ""
  }, {
    v1: "9.",
    v2: "Исправность средств сигнализации и контроля тормозных систем",
    v3: ""
  }, {
    v1: "10.",
    v2: "Отсутствие набухания тормозных шлангов под давлением, трещин и видимых мест перетирания",
    v3: ""
  }, {
    v1: "11.",
    v2: "Расположение и длина соединительных шлангов пневматического тормозного привода автопоездов",
    v3: ""
  }, {
    v1: "",
    v2: "Рулевое управление",
    v3: ""
  }, {
    v1: "12.",
    v2: "Работоспособность усилителя рулевого управления. Плавность изменения усилия при повороте рулевого колеса",
    v3: ""
  }, {
    v1: "13.",
    v2: "Отсутствие самопроизвольного поворота рулевого колеса с усилителем рулевого управления от нейтрального положения при работающем двигателе",
    v3: ""
  }, {
    v1: "14.",
    v2: "Отсутствие  превышения предельных значений суммарного люфта в рулевом управлении (не более 10º - легковые автомобили и созданные на базе их агрегатов грузовые автомобили и автобусы, 20º-автобусы, 25º-грузовые автомобили )",
    v3: ""
  }, {
    v1: "15.",
    v2: "Отсутствие повреждения и полная комплектность деталей крепления рулевой колонки и картера рулевого механизма",
    v3: ""
  }, {
    v1: "16.",
    v2: "Отсутствие следов остаточной деформации,  трещин и других дефектов в рулевом механизме и рулевом приводе",
    v3: ""
  }, {
    v1: "17.",
    v2: "Отсутствие устройств, ограничивающих поворот рулевого колеса, не предусмотренных конструкцией",
    v3: ""
  }, {
    v1: "39.",
    v2: "Соответствие норме светопропускания ветрового стекла (более 75%), передних боковых стекол и стекол передних дверей(более 70%)",
    v3: ""
  }, {
    v1: "",
    v2: "Двигатель и его системы",
    v3: ""
  }, {
    v1: "35.",
    v2: "Герметичность системы питания транспортных средств, работающих на газе. Соответствие газовых баллонов установленным требованиям",
    v3: ""
  }, {
    v1: "36.",
    v2: "Соответствие нормам уровня шума выпускной системы (не бьолее 96дБа - для М1,N1, 98дБа - для М2,N2, 100дБа - для М3,N3)",
    v3: ""
  }, {
    v1: "32.",
    v2: "Соответствие содержания загрязняющих веществ в отработавших газах транспортных средств установленным требованиям",
    v3: ""
  }, {
    v1: "Уровень СО и СН",
    v2: "Дымность отработанных газов, К  /м-1"
  }, {
    v1: "Наименование",
    v2: "Нормативные значения не больше",
    v3: "Результат",
    v4: "Зачетный замер",
    v5: "Результат измерений",
    v6: "Средний результат",
    v7: "Нормативное значение не больше"
  }, {
    v1: "число цилиндров",
    v2: "число цилиндров",
    v3: "1",
    v4: "2.5"
  }, {
    v1: "меньше 4",
    v2: "больше 4",
    v3: "меньше 4",
    v4: "больше 4",
    v5: "2"
  }, {
    v1: "СО,% на N мин",
    v2: "3.5",
    v3: "3"
  }, {
    v1: "СО,% на N увел.",
    v2: "2.0",
    v3: "4"
  }, {
    v1: "СН,млн-1 на N мин",
    v2: "1200",
    v3: "2500"
  }, {
    v1: "СН,млн-1 на N увел.",
    v2: "600",
    v3: "1000"
  }, {
    v1: "",
    v2: "Внешние световые приборы",
    v3: ""
  }, {
    v1: "18.",
    v2: "Соответствие устройств освещения и световой сигнализации  установленным требованиям",
    v3: ""
  }, {
    v1: "19.",
    v2: "Отсутствие разрушений  рассеивателей световых приборов",
    v3: ""
  }, {
    v1: "20.",
    v2: "Работоспособность и режим работы сигналов торможения",
    v3: ""
  }, {
    v1: "21.",
    v2: "Соответствие углов регулировки и силы света фар установленным требованиям",
    v3: ""
  }, {
    v1: "22.",
    v2: "Наличие и расположение фар и сигнальных фонарей в местах, предусмотренных конструкцией",
    v3: ""
  }, {
    v1: "Cоответствует",
    v2: "Не соответствует"
  }, {
    v1: "Дата прохождения",
    v2: "",
    v3: "Диагностическую карту выдал технический эксперт:",
    v4: ""
  }, {
    v1: "Подпись"
  }];

  //подставляем значения в ячейки
  ws.Cell(1, 1, 1, 6, true).Format.Font.Alignment.Horizontal('right').Format.Font.Bold().String('Диагностическая карта');
  ws.Cell(1, 7, 1, 10, true).Format.Font.Alignment.Horizontal('left').String('493467823612');
  ws.Cell(2, 1, 2, 10, true).Format.Font.Alignment.Horizontal('right').String('299053 г. Севастополь ул. Вакуленчука 41/9   ИП Пополитов Руслан Александрович');
  ws.Image('./123.png').Position(1, 11, 2, 12);

  for (var i = 3; i < 8; ++i) {
    ws.Cell(i, 1).String(data[i].v1).Style(myStyle);
    ws.Cell(i, 3, i, 4).String(data[i].v2).Style(myStyle2);
    ws.Cell(i, 5).String(data[i].v3).Style(myStyle);
    ws.Cell(i, 6, i, 7).String(data[i].v4).Style(myStyle2);
    ws.Cell(i, 8).String(data[i].v5).Style(myStyle);
    ws.Cell(i, 10, i, 12).String(data[i].v6).Style(myStyle2);
  }
  ws.Cell(8, 1).String(data[8].v1).Format.Font.Alignment.Horizontal('center');
  ws.Cell(8, 5, 8, 12).String(data[8].v2).Style(myStyle6);
  ws.Cell(9, 1).String(data[9].v1).Format.Font.Alignment.Horizontal('center');
  ws.Cell(9, 3, 9, 12).String(data[9].v2).Style(myStyle6);
  for (var i = 10; i < 12; ++i) {
    ws.Cell(i, 2).String(data[i].v1).Format.Font.Alignment.Horizontal('center');
    ws.Cell(i, 4).String(data[i].v2).Format.Font.Alignment.Horizontal('center');
    ws.Cell(i, 6).String(data[i].v3).Format.Font.Alignment.Horizontal('center');
    ws.Cell(i, 9).String(data[i].v4).Format.Font.Alignment.Horizontal('center');
  }
  for (var i = 12; i < 56; ++i) {
    ws.Cell(i, 1).String(data[i].v1).Style(myStyle4);
    ws.Cell(i, 2, i, 11).String(data[i].v2).Style(myStyle4);
    ws.Cell(i, 12).String(data[i].v3).Style(myStyle5);
  }
  ws.Cell(56, 2, 56, 8).String(data[56].v2);
  ws.Cell(56, 9).String(data[56].v3);
  ws.Cell(56, 10).String(data[56].v4).Style(myStyle4);
  ws.Cell(56, 11).String(data[56].v5).Style(myStyle4);
  ws.Cell(56, 12).String(data[56].v6).Style(myStyle4);
  for (var i = 57; i < 79; ++i) {
    ws.Cell(i, 1).String(data[i].v1).Style(myStyle4);
    ws.Cell(i, 2, i, 11).String(data[i].v2).Style(myStyle4);
    ws.Cell(i, 12).String(data[i].v3).Style(myStyle5);
  }
  ws.Cell(79, 1).String(data[79].v1).Style(myStyle5);
  ws.Cell(79, 8).String(data[79].v2).Style(myStyle5);
  ws.Cell(80, 1).String(data[80].v1).Style(myStyle5);
  ws.Cell(80, 3).String(data[80].v2).Style(myStyle5);
  ws.Cell(80, 5).String(data[80].v3).Style(myStyle5);
  ws.Cell(80, 8).String(data[80].v4).Style(myStyle5);
  ws.Cell(80, 9).String(data[80].v5).Style(myStyle5);
  ws.Cell(80, 10).String(data[80].v6).Style(myStyle5);
  ws.Cell(80, 11).String(data[80].v7).Style(myStyle5);
  ws.Cell(81, 3).String(data[81].v1).Style(myStyle5);
  ws.Cell(81, 5).String(data[81].v2).Style(myStyle5);
  ws.Cell(81, 8).String(data[81].v3).Style(myStyle5);
  ws.Cell(81, 11).String(data[81].v4).Style(myStyle5);
  ws.Cell(82, 4).String(data[82].v2).Style(myStyle5);
  ws.Cell(82, 3).String(data[82].v1).Style(myStyle5);
  ws.Cell(82, 5).String(data[82].v3).Style(myStyle5);
  ws.Cell(82, 6).String(data[82].v4).Style(myStyle5);
  ws.Cell(82, 8).String(data[82].v5).Style(myStyle5);
  ws.Cell(83, 1).String(data[83].v1).Style(myStyle5);
  ws.Cell(83, 3).String(data[83].v2).Style(myStyle5);
  ws.Cell(83, 8).String(data[83].v3).Style(myStyle5);
  ws.Cell(84, 1).String(data[84].v1).Style(myStyle5);
  ws.Cell(84, 3).String(data[84].v2).Style(myStyle5);
  ws.Cell(84, 8).String(data[84].v3).Style(myStyle5);
  ws.Cell(85, 1).String(data[85].v1).Style(myStyle5);
  ws.Cell(85, 3).String(data[85].v2).Style(myStyle5);
  ws.Cell(85, 4).String(data[85].v3).Style(myStyle5);
  ws.Cell(86, 1).String(data[86].v1).Style(myStyle5);
  ws.Cell(86, 3).String(data[86].v2).Style(myStyle5);
  ws.Cell(86, 4).String(data[86].v3).Style(myStyle5);
  ws.Cell(87, 2).String(data[87].v1).Style(myStyle5);
  for (var i = 87; i < 93; ++i) {
    ws.Cell(i, 1).String(data[i].v1).Style(myStyle4);
    ws.Cell(i, 2).String(data[i].v2).Style(myStyle4);
    ws.Cell(i, 12).String(data[i].v3).Style(myStyle5);
  }
  ws.Cell(93, 8).String(data[93].v1);
  ws.Cell(93, 10).String(data[93].v2);
  ws.Cell(94, 1).String(data[94].v1);
  ws.Cell(94, 4, 94, 5).String(data[94].v2).Style(myStyle2);
  ws.Cell(94, 6).String(data[94].v3);
  ws.Cell(94, 10, 94, 12).String(data[94].v4).Style(myStyle2);
  ws.Cell(95, 9).String(data[95].v1);

  //formatting cells in yellow
  ws.Cell(13, 12, 16, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(24, 12, 26, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(37, 12, 38, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(40, 12, 42, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(44, 12, 46, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(48, 12, 53, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(55, 12, 56, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(58, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(60, 12, 63, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(65, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(68, 12, 74, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(77, 12, 78, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');
  ws.Cell(88, 12, 92, 12).Format.Fill.Color('FFF000').Format.Fill.Pattern('solid');

  //formatting cells in Bold
  ws.Cell(43, 2).Style(myStyle4Bold);
  ws.Cell(46, 2).Style(myStyle4Bold);
  ws.Cell(47, 2).Style(myStyle4Bold);
  ws.Cell(48, 2).Style(myStyle4Bold);
  ws.Cell(54, 2).Style(myStyle4Bold);
  ws.Cell(55, 2).Style(myStyle4Bold);
  ws.Cell(56, 2).Style(myStyle4Bold);
  ws.Cell(58, 2).Style(myStyle4Bold);
  ws.Cell(67, 2).Style(myStyle4Bold);
  ws.Cell(70, 2).Style(myStyle4Bold);
  ws.Cell(75, 2).Style(myStyle4Bold);
  ws.Cell(77, 2).Style(myStyle4Bold);
  ws.Cell(87, 2).Style(myStyle4Bold);

  //Border
  ws.Cell(1, 13, 95, 13).Style(leftBorder);
  ws.Cell(95, 1, 95, 12).Style(bottomBorder);

  return data;
}

function createFirstSheet() {
  var XLSX = require('xlsx');
  var workbook = XLSX.readFile(path.normalize(journalPath));
  var firstSheetName = workbook.SheetNames[0];
  var worksheet = workbook.Sheets[firstSheetName];
  var array = XLSX.utils.sheet_to_json(worksheet);
  var myStyle = journal.Style();
  myStyle.Font.Alignment.Vertical('center');
  myStyle.Font.Alignment.Horizontal('center');
  myStyle.Font.Family('Times New Roman');
  myStyle.Font.Size(10);
  myStyle.Font.WrapText();
  var blackStyle = journal.Style();
  blackStyle.Font.Size(12);
  blackStyle.Font.Family('Times New Roman');
  blackStyle.Font.Alignment.Vertical('center');
  blackStyle.Font.Alignment.Horizontal('center');
  blackStyle.Font.Color('FFFFFF');
  blackStyle.Font.WrapText();
  blackStyle.Fill.Color('000000');
  blackStyle.Fill.Pattern('solid');
  /*
    Code to generate page Журнал регистрации ТС
  */
  ws0.Column(1).Width(25);
  ws0.Column(3).Width(20);
  ws0.Column(4).Width(20);
  ws0.Column(6).Width(20);
  ws0.Column(20).Width(30);

  ws0.Cell(1, 1).String('РЕГИСТРАЦИОННЫЙ НОМЕР').Style(blackStyle);
  ws0.Cell(1, 2).String('Владелец ТС ').Style(blackStyle);
  ws0.Cell(1, 3).String('Регестрационный знак ТС').Style(blackStyle);
  ws0.Cell(1, 4).String('VIN').Style(blackStyle);
  ws0.Cell(1, 5).String('№ шасси, рамы').Style(blackStyle);
  ws0.Cell(1, 6).String('№ кузова').Style(blackStyle);
  ws0.Cell(1, 7).String('Марка ТС').Style(blackStyle);
  ws0.Cell(1, 8).String('Модель ТС').Style(blackStyle);
  ws0.Cell(1, 9).String('Колич. авто').Style(blackStyle);
  ws0.Cell(1, 10).String('Категория ТС').Style(blackStyle);
  ws0.Cell(1, 11).String('Год выпуска ТС').Style(blackStyle);
  ws0.Cell(1, 12).String('Марка шин').Style(blackStyle);
  ws0.Cell(1, 13).String('Пробег ТС').Style(blackStyle);
  ws0.Cell(1, 14).String('Тип топлива').Style(blackStyle);
  ws0.Cell(1, 15).String('Тип тормозной системы').Style(blackStyle);
  ws0.Cell(1, 16).String('Разрешенная мах  масса').Style(blackStyle);
  ws0.Cell(1, 17).String('Масса без нагрузки').Style(blackStyle);
  ws0.Cell(1, 18).String('СРТС или ПТС серия').Style(blackStyle);
  ws0.Cell(1, 19).String('СРТС или ПТС  номер').Style(blackStyle);
  ws0.Cell(1, 20).String('СРТС или ПТС кем выдан кем').Style(blackStyle);
  ws0.Cell(1, 21).String('СРТС или ПТС  когда выдан').Style(blackStyle);
  ws0.Cell(1, 22).String('Дата ТК').Style(blackStyle);
  ws0.Cell(1, 23).String('Срок действия до: ').Style(blackStyle);
  ws0.Cell(1, 24).String('Ф.И.О. технического эксперта ').Style(blackStyle);
  ws0.Cell(1, 25).String('Цена').Style(blackStyle);
  ws0.Cell(1, 26).String('Письменно').Style(blackStyle);
  ws0.Cell(1, 27).String('ЕАИСТО').Style(blackStyle);
  ws0.Cell(1, 28).String('Договор').Style(blackStyle);
  ws0.Cell(1, 29).String('ТЕЛ').Style(blackStyle);
  ws0.Cell(1, 30).String('№ Счета').Style(blackStyle);

  console.log('Reading File...');
  for (var i = 0; i < array.length; i++) {
    var j = 1;
    for (key in array[i]) {
      ws0.Cell(i + 2, j).String(array[i][key]).Style(myStyle);
      j++;
    }
  }
  console.log('Reading file has been completed!');
  regNumber = (parseInt(array[array.length - 1]['РЕГИСТРАЦИОННЫЙ НОМЕР']) + 1).toString();
  if (regNumber.length < 15)
    regNumber = '0' + regNumber;
  ws0.Cell(array.length + 2, 1).String(regNumber).Style(myStyle);
  ws0.Cell(2, 2).String(post.secondName + " " + post.firstName + " " + post.lastName).Style(myStyle);
  ws0.Cell(2, 3).String(post.regPlate).Style(myStyle);
  ws0.Cell(2, 4).String(post.vin).Style(myStyle);
  ws0.Cell(2, 5).String(post.chassis).Style(myStyle);
  ws0.Cell(2, 6).String(post.body).Style(myStyle);
  ws0.Cell(2, 7).String(post.mark).Style(myStyle);
  ws0.Cell(2, 8).String(post.model).Style(myStyle);
  ws0.Cell(2, 9).String('1').Style(myStyle);
  ws0.Cell(2, 10).String(post.vehicleCategory).Style(myStyle);
  ws0.Cell(2, 11).String(post.year).Style(myStyle);
  ws0.Cell(2, 12).String(post.tire).Style(myStyle);
  ws0.Cell(2, 13).String(post.mileage).Style(myStyle);
  ws0.Cell(2, 14).String(post.fuel).Style(myStyle);
  ws0.Cell(2, 15).String(post.brakeSystem).Style(myStyle);
  ws0.Cell(2, 16).String(post.maxWeight).Style(myStyle);
  ws0.Cell(2, 17).String(post.weight).Style(myStyle);
  ws0.Cell(2, 18).String(post.regSERIA).Style(myStyle);
  ws0.Cell(2, 19).String(post.regNOMER).Style(myStyle);
  ws0.Cell(2, 20).String(post.regKEM).Style(myStyle);
  ws0.Cell(2, 21).String(post.regKOGDA).Style(myStyle);
  ws0.Cell(2, 22).String('').Style(myStyle);
  ws0.Cell(2, 23).String('').Style(myStyle);
  ws0.Cell(2, 24).String('').Style(myStyle);
  ws0.Cell(2, 25).String('').Style(myStyle);
  ws0.Cell(2, 26).String('').Style(myStyle);
  ws0.Cell(2, 27).String('').Style(myStyle);
  ws0.Cell(2, 28).String('').Style(myStyle);
  ws0.Cell(2, 29).String('').Style(myStyle);
  ws0.Cell(2, 30).String('').Style(myStyle);
}

function createSecondSheet(data1, post) {
  ws2 = wb.WorkSheet('ДК лист 1', wsOpts);
  var data = [{}, {
    v1: ""
  }, {
    v1: ""
  }, {
    v1: ""
  }, {
    v1: ""
  }, {
    v1: 'Первичная проверка',
    v2: '', //wtf
    v3: 'Повторная проверка',
    v4: '' //wtf
  }, {
    v1: 'Регистрационный знак ТС: ',
    v2: data1[4].v2,
    v3: 'Марка, модель ТС:',
    v4: data1[4].v4
  }, {
    v1: 'VIN',
    v2: data1[5].v2,
    v3: 'Категория ТС:',
    v4: data1[5].v4
  }, {
    v1: 'Номер шасси, рамы:',
    v2: data1[6].v2,
    v3: 'Год выпуска ТС:',
    v4: data1[6].v4
  }, {
    v1: 'Номер кузова:',
    v2: data1[7].v2,
    v3: '',
    v4: ''
  }, {
    v1: 'СРТС или ПТС (серия, номер, выдан кем, когда):',
    v2: data1[8].v2
  }, {
    v1: '№',
    v2: 'Параметры и требования, предъявляемые к транспортным средствам при проведении технического контроля',
    v3: '',
    v4: '№',
    v5: 'Параметры и требования, предъявляемые к транспортным средствам при проведении технического контроля',
    v6: '',
    v7: '№',
    v8: 'Параметры и требования, предъявляемые к транспортным средствам при проведении технического контроля',
    v9: ''
  }, {
    v1: '',
    v2: 'I. Тормозные системы',
    v3: '',
    v4: '22.',
    v5: 'Наличие и расположение фар и сигнальных фонарей в местах, предусмотренных конструкцией',
    v6: '',
    v7: '42.',
    v8: 'Работоспособность запоров бортов грузовой платформы и запоров горловин цистерн',
    v9: ''
  }, {
    v1: '1',
    v2: 'Соответствие показателей эффективности торможения и устойчивости торможения',
    v3: '',
    v4: '',
    v5: 'IV. Стеклоочистители и стеклоомыватели',
    v6: '',
    v7: '43.',
    v8: 'Работоспособность аварийного выключателя дверей и сигнала требования остановки',
    v9: ''
  }, {
    v1: '2.',
    v2: 'Соответствие разности тормозных сил установленным требованиям',
    v3: '',
    v4: '23.',
    v5: 'Наличие  стеклоочистителя и форсунки стеклоомывателя ветрового стекла',
    v6: '',
    v7: '44.',
    v8: 'Работоспособность аварийных выходов, приборов внутреннего освещения салона, привода управления дверями и сигнализации их работы',
    v9: ''
  }, {
    v1: '3.',
    v2: 'Работоспособность рабочей тормозной системы автопоездов с пневматическим тормозным приводом в режиме аварийного (автоматического) торможения',
    v3: '',
    v4: '24.',
    v5: 'Обеспечение стеклоомывателем подачи жидкости в зоны очистки стекла',
    v6: '',
    v7: '45.',
    v8: 'Наличие работоспособного звукового сигнального прибора',
    v9: ''
  }, {
    v1: '4.',
    v2: 'Отсутствие утечек сжатого воздуха из колесных тормозных камер',
    v3: '',
    v4: '25.',
    v5: 'Работоспособность стеклоочистителей и стеклоомывателей',
    v6: '',
    v7: '46.',
    v8: 'Наличие обозначений аварийных выходов и табличек по правилам их использования. Обеспечение свободного доступа к аварийным выходам',
    v9: ''
  }, {
    v1: '5.',
    v2: 'Отсутствие подтеканий тормозной жидкости, нарушения герметичности трубопроводов или соединений в гидравлическом тормозном приводе',
    v3: '',
    v4: '',
    v5: 'V. Шины и колеса',
    v6: '',
    v7: '47.',
    v8: 'Наличие задних и боковых  защитных устройств, соответствие их нормам',
    v9: ''
  }, {
    v1: '6.',
    v2: 'Отсутствие коррозии, грозящей потерей герметичности или разрушением',
    v3: '',
    v4: '26.',
    v5: 'Соответствие высоты рисунка протектора шин установленным требованиям',
    v6: '',
    v7: '48.',
    v8: 'Работоспособность автоматического замка, ручной и автоматической блокировки седельно-сцепного устройства. Отсутствие видимых повреждений сцепных устройств',
    v9: ''
  }, {
    v1: '7.',
    v2: 'Отсутствие механических повреждений тормозных трубопроводов',
    v3: '',
    v4: '27.',
    v5: 'Отсутствие признаков непригодности шин к эксплуатации',
    v6: '',
    v7: '49.',
    v8: 'Наличие работоспособных предохранительных приспособлений у одноосных прицепов (за исключением роспусков) и прицепов, не оборудованных рабочей тормозной системой',
    v9: ''
  }, {
    v1: '8.',
    v2: 'Отсутствие трещин остаточной деформации деталей тормозного привода',
    v3: '',
    v4: '28.',
    v5: 'Наличие всех болтов или гаек крепления дисков и ободьев колес',
    v6: '',
    v7: '50.',
    v8: 'Оборудование прицепов (за исключением одноосных и роспусков) исправным устройством, поддерживающим сцепную петлю дышла в положении, облегчающем сцепку и расцепку с тяговым автомобилем',
    v9: ''
  }, {
    v1: '9.',
    v2: 'Исправность средств сигнализации и контроля тормозных систем',
    v3: '',
    v4: '29.',
    v5: 'Отсутствие трещин на дисках и ободьях колес',
    v6: '',
    v7: '51.',
    v8: 'Отсутствие продольного люфта в беззазорных тягово-сцепных устройствах с тяговой вилкой для сцепленного с прицепом тягача',
    v9: ''
  }, {
    v1: '10.',
    v2: 'Отсутствие набухания тормозных шлангов под давлением, трещин и видимых мест перетирания',
    v3: '',
    v4: '30.',
    v5: 'Отсутствие видимых нарушений формы и размеров крепежных отверстий в дисках колес',
    v6: '',
    v7: '52.',
    v8: 'Обеспечение тягово-сцепными устройствами легковых автомобилей беззазорной сцепки сухарей замкового устройства с шаром',
    v9: ''
  }, {
    v1: '11.',
    v2: 'Расположение и длина соединительных шлангов пневматического тормозного привода автопоездов',
    v3: '',
    v4: '31.',
    v5: 'Установка шин на транспортное средство в соответствии с требованиями',
    v6: '',
    v7: '53.',
    v8: 'Соответствие размерных характеристик сцепных устройств установленным требованиям',
    v9: ''
  }, {
    v1: '',
    v2: 'II. Рулевое управление',
    v3: '',
    v4: '',
    v5: 'VI. Двигатель и его системы',
    v6: '',
    v7: '54.',
    v8: 'Оснащение транспортных средств исправными ремнями безопасности',
    v9: ''
  }, {
    v1: '12.',
    v2: 'Работоспособность усилителя рулевого управления. Плавность изменения усилия при повороте рулевого колеса',
    v3: '',
    v4: '32.',
    v5: 'Соответствие содержания загрязняющих веществ в отработавших газах транспортных средств установленным требованиям',
    v6: '',
    v7: '55.',
    v8: 'Наличие знака аварийной остановки',
    v9: ''
  }, {
    v1: '13.',
    v2: 'Отсутствие самопроизвольного поворота рулевого колеса с усилителем рулевого управления от нейтрального положения при работающем двигателе',
    v3: '',
    v4: '33.',
    v5: 'Отсутствие подтекания и каплепадения топлива в системе питания',
    v6: '',
    v7: '56.',
    v8: 'Наличие не менее двух противооткатных упоров',
    v9: ''
  }, {
    v1: '14.',
    v2: 'Отсутствие  превышения предельных значений суммарного люфта в рулевом управлении',
    v3: '',
    v4: '34.',
    v5: 'Работоспособность запорных устройств и устройств перекрытия топлива',
    v6: '',
    v7: '57.',
    v8: 'Наличие огнетушителей, соответствующих установленным требованиям',
    v9: ''
  }, {
    v1: '15.',
    v2: 'Отсутствие повреждения и полная комплектность деталей крепления рулевой колонки и картера рулевого механизма',
    v3: '',
    v4: '35.',
    v5: 'Герметичность системы питания транспортных средств, работающих на газе. Соответствие газовых баллонов установленным требованиям',
    v6: '',
    v7: '58.',
    v8: 'Надежное крепление поручней в автобусах, запасного колеса, аккумуляторной батареи, сидений, огнетушителей и медицинской аптечки',
    v9: ''
  }, {
    v1: '16.',
    v2: 'Отсутствие следов остаточной деформации,  трещин и других дефектов в рулевом механизме и рулевом приводе',
    v3: '',
    v4: '36.',
    v5: 'Соответствие нормам уровня шума выпускной системы',
    v6: '',
    v7: '59.',
    v8: 'Работоспособность механизмов регулировки сидений',
    v9: ''
  }, {
    v1: '17.',
    v2: 'Отсутствие устройств, ограничивающих поворот рулевого колеса, не предусмотренных конструкцией',
    v3: '',
    v4: '',
    v5: 'VII. Прочие элементы конструкции',
    v6: '',
    v7: '60.',
    v8: 'Наличие надколесных грязезащитных устройств, отвечающих установленным требованиям',
    v9: ''
  }, {
    v1: '',
    v2: 'III. Внешние световые приборы',
    v3: '',
    v4: '37.',
    v5: 'Наличие зеркал заднего вида в соответствии с требованиями',
    v6: '',
    v7: '61.',
    v8: 'Соответствие вертикальной статической нагрузки на тяговое устройство автомобиля от сцепной петли одноосного прицепа (прицепа-роспуска) нормам',
    v9: ''
  }, {
    v1: '18.',
    v2: 'Соответствие устройств освещения и световой сигнализации  установленным требованиям',
    v3: '',
    v4: '38.',
    v5: 'Отсутствие дополнительных предметов или покрытий, ограничивающих обзорность с места водителя. Соответствие полосы пленки в  верхней  части ветрового  стекла  установленным требованиям',
    v6: '',
    v7: '62.',
    v8: 'Работоспособность держателя запасного колеса, лебедки и механизма подъема-опускания запасного колеса',
    v9: ''
  }, {
    v1: '19.',
    v2: 'Отсутствие разрушений  рассеивателей световых приборов',
    v3: '',
    v4: '39.',
    v5: 'Соответствие норме светопропускания ветрового стекла, передних боковых стекол и стекол передних дверей',
    v6: '',
    v7: '63.',
    v8: 'Работоспособность механизмов подъема и опускания опор и фиксаторов транспортного положения опор',
    v9: ''
  }, {
    v1: '20.',
    v2: 'Работоспособность и режим работы сигналов торможения',
    v3: '',
    v4: '40.',
    v5: 'Отсутствие трещин на ветровом стекле в зоне очистки водительского стеклоочистителя',
    v6: '',
    v7: '64.',
    v8: 'Соответствие каплепадения масел и рабочих жидкостей нормам',
    v9: ''
  }, {
    v1: '21.',
    v2: 'Соответствие углов регулировки и силы света фар установленным требованиям',
    v3: '',
    v4: '41.',
    v5: 'Работоспособность замков дверей кузова,  кабины, механизмов регулировки и фиксирующих устройств сидений, устройства обогрева и обдува ветрового стекла, противоугонного устройства',
    v6: '',
    v7: '65.',
    v8: 'Установка государственных регистрационных знаков в соответствии с требованиями',
    v9: ''
  }];
  var myStyle = wb.Style();
  myStyle.Font.Size(10);
  myStyle.Font.Family('Times New Roman');
  myStyle.Font.Alignment.Vertical('center');
  myStyle.Font.Alignment.Horizontal('center');
  myStyle.Font.WrapText();
  var myStyle2 = wb.Style();
  myStyle2.Fill.Color('C5F8FF');
  myStyle2.Fill.Pattern('solid');
  myStyle2.Font.Family('Times New Roman');
  myStyle2.Font.Alignment.Vertical('center');
  myStyle2.Font.Alignment.Horizontal('center');
  myStyle2.Font.WrapText();
  myStyle2.Border({
    left: {
      style: 'thin',
      color: 'D0D0D0'
    },
    right: {
      style: 'thin',
      color: 'D0D0D0'
    },
    top: {
      style: 'thin',
      color: 'D0D0D0'
    },
    bottom: {
      style: 'thin',
      color: 'D0D0D0'
    }
  });
  var leftBorder = wb.Style();
  leftBorder.Border({
    left: {
      style: 'thick',
      color: '0000FF'
    }
  });
  var bottomBorder = wb.Style();
  bottomBorder.Border({
    bottom: {
      style: 'thick',
      color: '0000FF'
    }
  });

  ws2.Column(1).Width(4);
  ws2.Column(2).Width(10);
  ws2.Column(3).Width(0);
  ws2.Column(4).Width(10);
  ws2.Column(5).Width(10);
  ws2.Column(6).Width(3);
  ws2.Column(7).Width(4);
  ws2.Column(8).Width(15);
  ws2.Column(9).Width(12);
  ws2.Column(10).Width(0);
  ws2.Column(11).Width(3);
  ws2.Column(12).Width(4);
  ws2.Column(13).Width(15);
  ws2.Column(14).Width(15);
  ws2.Column(15).Width(3);

  //объединение ячеек
  ws2.Cell(1, 1, 1, 15, true).Format.Font.Alignment.Horizontal('center').String('Диагностическая карта  Certificate of periodic technical inspection');
  ws2.Cell(2, 1, 2, 4, true).Format.Font.Alignment.Horizontal('left').String('Регистрационный номер');
  ws2.Cell(2, 5, 2, 7, true);
  ws2.Cell(2, 12, 2, 13, true).Format.Font.Alignment.Horizontal('left').Format.Font.Bold().String('Срок действия до:');
  ws2.Cell(2, 14, 2, 15, true);
  ws2.Cell(3, 1, 3, 5, true).String('Оператор технического осмотра:');
  ws2.Cell(3, 6, 3, 15, true);
  ws2.Cell(4, 1, 4, 5, true).String('Пункт технического осмотра:');
  ws2.Cell(4, 6, 4, 15, true);

  for (var i = 5; i < 10; ++i) {
    ws2.Cell(i, 1, i, 4, true).String(data[i].v1);
    ws2.Cell(i, 5, i, 8, true).String(data[i].v2).Style(myStyle2);
    ws2.Cell(i, 9, i, 12, true).String(data[i].v3);
    ws2.Cell(i, 13, i, 15, true).String(data[i].v4).Style(myStyle2);
  }

  ws2.Cell(10, 1, 10, 6, true).String(data[10].v1).Style(myStyle);
  ws2.Cell(10, 7, 10, 15, true).String(data[10].v2).Style(myStyle2);

  //data for first column
  ws2.Cell(13, 6).String(post._1).Style(myStyle2);
  ws2.Cell(14, 6).String(post._2).Style(myStyle2);
  ws2.Cell(15, 6).String(post._3).Style(myStyle2);
  ws2.Cell(16, 6).String(post._4).Style(myStyle2);
  ws2.Cell(17, 6).String(post._5).Style(myStyle2);
  ws2.Cell(18, 6).String(post._6).Style(myStyle2);
  ws2.Cell(19, 6).String(post._7).Style(myStyle2);
  ws2.Cell(20, 6).String(post._8).Style(myStyle2);
  ws2.Cell(21, 6).String(post._9).Style(myStyle2);
  ws2.Cell(22, 6).String(post._10).Style(myStyle2);
  ws2.Cell(23, 6).String(post._11).Style(myStyle2);
  ws2.Cell(25, 6).String(post._12).Style(myStyle2);
  ws2.Cell(26, 6).String(post._13).Style(myStyle2);
  ws2.Cell(27, 6).String(post._14).Style(myStyle2);
  ws2.Cell(28, 6).String(post._15).Style(myStyle2);
  ws2.Cell(29, 6).String(post._16).Style(myStyle2);
  ws2.Cell(30, 6).String(post._17).Style(myStyle2);
  ws2.Cell(32, 6).String(post._18).Style(myStyle2);
  ws2.Cell(33, 6).String(post._19).Style(myStyle2);
  ws2.Cell(34, 6).String(post._20).Style(myStyle2);
  ws2.Cell(35, 6).String(post._21).Style(myStyle2);
  //data for second column
  ws2.Cell(12, 11).String(post._22).Style(myStyle2);
  ws2.Cell(14, 11).String(post._23).Style(myStyle2);
  ws2.Cell(15, 11).String(post._24).Style(myStyle2);
  ws2.Cell(16, 11).String(post._25).Style(myStyle2);
  ws2.Cell(18, 11).String(post._26).Style(myStyle2);
  ws2.Cell(19, 11).String(post._27).Style(myStyle2);
  ws2.Cell(20, 11).String(post._28).Style(myStyle2);
  ws2.Cell(21, 11).String(post._29).Style(myStyle2);
  ws2.Cell(22, 11).String(post._30).Style(myStyle2);
  ws2.Cell(23, 11).String(post._31).Style(myStyle2);
  ws2.Cell(25, 11).String(post._32).Style(myStyle2);
  ws2.Cell(26, 11).String(post._33).Style(myStyle2);
  ws2.Cell(27, 11).String(post._34).Style(myStyle2);
  ws2.Cell(28, 11).String(post._35).Style(myStyle2);
  ws2.Cell(29, 11).String(post._36).Style(myStyle2);
  ws2.Cell(31, 11).String(post._37).Style(myStyle2);
  ws2.Cell(32, 11).String(post._38).Style(myStyle2);
  ws2.Cell(33, 11).String(post._39).Style(myStyle2);
  ws2.Cell(34, 11).String(post._40).Style(myStyle2);
  ws2.Cell(35, 11).String(post._41).Style(myStyle2);
  //data for third column
  ws2.Cell(12, 15).String(post._42).Style(myStyle2);
  ws2.Cell(13, 15).String(post._43).Style(myStyle2);
  ws2.Cell(14, 15).String(post._44).Style(myStyle2);
  ws2.Cell(15, 15).String(post._45).Style(myStyle2);
  ws2.Cell(16, 15).String(post._46).Style(myStyle2);
  ws2.Cell(17, 15).String(post._47).Style(myStyle2);
  ws2.Cell(18, 15).String(post._48).Style(myStyle2);
  ws2.Cell(19, 15).String(post._49).Style(myStyle2);
  ws2.Cell(20, 15).String(post._50).Style(myStyle2);
  ws2.Cell(21, 15).String(post._51).Style(myStyle2);
  ws2.Cell(22, 15).String(post._52).Style(myStyle2);
  ws2.Cell(23, 15).String(post._53).Style(myStyle2);
  ws2.Cell(24, 15).String(post._54).Style(myStyle2);
  ws2.Cell(25, 15).String(post._55).Style(myStyle2);
  ws2.Cell(26, 15).String(post._56).Style(myStyle2);
  ws2.Cell(27, 15).String(post._57).Style(myStyle2);
  ws2.Cell(28, 15).String(post._58).Style(myStyle2);
  ws2.Cell(29, 15).String(post._59).Style(myStyle2);
  ws2.Cell(30, 15).String(post._60).Style(myStyle2);
  ws2.Cell(31, 15).String(post._61).Style(myStyle2);
  ws2.Cell(32, 15).String(post._62).Style(myStyle2);
  ws2.Cell(33, 15).String(post._63).Style(myStyle2);
  ws2.Cell(34, 15).String(post._64).Style(myStyle2);
  ws2.Cell(35, 15).String(post._65).Style(myStyle2);

  //opt
  for (var i = 11; i < 36; i++) {
    ws2.Row(i).Height(50);
    ws2.Cell(i, 1).String(data[i].v1).Style(myStyle);
    ws2.Cell(i, 2, i, 5, true).String(data[i].v2).Style(myStyle);
    ws2.Cell(i, 7).String(data[i].v4).Style(myStyle);
    ws2.Cell(i, 8, i, 9, true).String(data[i].v5).Style(myStyle);
    ws2.Cell(i, 12).String(data[i].v7).Style(myStyle);
    ws2.Cell(i, 13, i, 14, true).String(data[i].v8).Style(myStyle);
  }
  //Border
  ws2.Cell(1, 16, 35, 16).Style(leftBorder);
  ws2.Cell(35, 1, 35, 15).Style(bottomBorder);
}


function createThirdSheet(data1, post) {
  ws3 = wb.WorkSheet('ДК лист 2', wsOpts);

  ws3.Column(1).Width(10);
  ws3.Column(2).Width(12);
  ws3.Column(3).Width(10);
  ws3.Column(4).Width(20);
  ws3.Column(5).Width(15);
  ws3.Column(6).Width(15);
  ws3.Column(7).Width(17);

  ws3.Row(48).Height(30);
  //стили
  var myStyle = wb.Style();
  myStyle.Font.Size(12);
  myStyle.Font.Family('Times New Roman');
  myStyle.Font.Alignment.Vertical('center');
  myStyle.Font.Alignment.Horizontal('center');
  myStyle.Font.WrapText();

  var myStyle2 = wb.Style();
  myStyle2.Fill.Color('C5F8FF');
  myStyle2.Fill.Pattern('solid');
  myStyle2.Font.Size(12);
  myStyle2.Font.Family('Times New Roman');
  myStyle2.Font.Alignment.Vertical('center');
  myStyle2.Font.Alignment.Horizontal('center');
  myStyle2.Font.WrapText();
  myStyle2.Border({
    left: {
      style: 'thin',
      color: 'D0D0D0'
    },
    right: {
      style: 'thin',
      color: 'D0D0D0'
    },
    top: {
      style: 'thin',
      color: 'D0D0D0'
    },
    bottom: {
      style: 'thin',
      color: 'D0D0D0'
    }
  });
  var myStyle3 = wb.Style();
  myStyle3.Font.Size(12);
  myStyle3.Font.Family('Times New Roman');
  myStyle3.Font.Alignment.Vertical('center');
  myStyle3.Font.Alignment.Horizontal('left');
  myStyle3.Font.WrapText();

  var myStyle4 = wb.Style();
  myStyle4.Font.Size(11);
  myStyle4.Font.Family('Times New Roman');
  myStyle4.Font.WrapText();

  var leftBorder = wb.Style();
  leftBorder.Border({
    left: {
      style: 'thick',
      color: '0000FF'
    }
  });
  var bottomBorder = wb.Style();
  bottomBorder.Border({
    bottom: {
      style: 'thick',
      color: '0000FF'
    }
  });
  //Журнал регистрации
  ws0.Cell(2, 22).String(post.date);
  ws0.Cell(2, 23).String(post.validity);
  ws0.Cell(2, 24).String(post.expert);
  //Первый лист
  if (post.checkType == '1')
    ws.Cell(3, 3).String('X');
  else if (post.checkType == '2')
    ws.Cell(3, 6).String('X');
  if (post.result == '1')
    ws.Cell(93, 8, 93, 9).Style(myStyle2);
  else if (post.result == '2')
    ws.Cell(93, 10, 93, 11).Style(myStyle2);
  ws.Cell(94, 4).String(post.date);
  ws.Cell(94, 10).String(post.expert);
  //Второй лист срок действия
  ws2.Cell(2, 14).String(post.validity);
  if (post.checkType == '1') {
    ws2.Cell(5, 5).String('X');
  } else if (post.checkType == '2') {
    ws2.Cell(5, 13).String('X');
  }
  ws2.Cell(2, 14).String(post.validity);
  ws2.Cell(3, 6).String(post.expert);
  //объединение ячеек
  ws3.Cell(1, 1, 1, 7, true).String('Результаты диагностирования').Style(myStyle);
  ws3.Cell(2, 1, 2, 6, true).String('Параметры, по которым установлено несоответствие').Style(myStyle);
  ws3.Cell(2, 7, 3, 7, true).String('Пункт диагностической карты').Style(myStyle);
  ws3.Cell(3, 1).String('Нижняя граница').Style(myStyle);
  ws3.Cell(3, 2).String('Результат проверки').Style(myStyle);
  ws3.Cell(3, 3).String('Верхняя граница').Style(myStyle);
  ws3.Cell(3, 4, 3, 6, true).String('Наименование параметра').Style(myStyle);
  ws3.Cell(3, 7).String('Пункт диагностической карты').Style(myStyle);
  for (var i = 4; i < 14; ++i) {
    ws3.Cell(i, 1).Style(myStyle);
    ws3.Cell(i, 2).Style(myStyle);
    ws3.Cell(i, 3).Style(myStyle);
    ws3.Cell(i, 4, i, 6, true).Style(myStyle);
    ws3.Cell(i, 7).Style(myStyle);
  }
  for (var i = 0; i < diagnosticIssues.length; ++i) {
    ws3.Cell(i + 4, 1).String(diagnosticIssues[i].v1);
    ws3.Cell(i + 4, 2).String(diagnosticIssues[i].v2);
    ws3.Cell(i + 4, 3).String(diagnosticIssues[i].v3);
    ws3.Cell(i + 4, 4).String(diagnosticIssues[i].v4);
    ws3.Cell(i + 4, 7).String(diagnosticIssues[i].v5);
  }
  ws3.Cell(14, 1, 14, 7, true).String('Невыполненные требования').Style(myStyle);
  ws3.Cell(15, 1, 15, 2, true).String('Предмет проверки (узел, деталь, агрегат)').Style(myStyle);
  ws3.Cell(15, 3, 15, 6, true).String('Содержание невыполненного требования (с указанием нормативного источника)').Style(myStyle);
  ws3.Cell(15, 7).String('Пункт диагностической карты').Style(myStyle);
  for (var i = 16; i < 27; ++i) {
    ws3.Cell(i, 1, i, 2, true).Style(myStyle);
    ws3.Cell(i, 3, i, 6, true).Style(myStyle);
    ws3.Cell(i, 7).Style(myStyle);
  }
  for (var i = 0; i < failedRequirements.length; ++i) {
    ws3.Cell(i + 16, 1).String(failedRequirements[i].v1);
    ws3.Cell(i + 16, 3).String(failedRequirements[i].v2);
    ws3.Cell(i + 16, 7).String(failedRequirements[i].v3);
  }
  ws3.Cell(27, 1, 27, 7, true).Format.Font.Alignment.Horizontal('left').String('Примечания:');
  ws3.Cell(28, 1, 32, 7, true).String(notes).Style(myStyle3);
  ws3.Cell(33, 1, 33, 7, true).String('Данные транспортного средства').Style(myStyle);
  ws3.Cell(34, 1, 34, 2, true).String(data1[7].v5).Style(myStyle3);
  ws3.Cell(34, 3, 34, 4, true).String(data1[7].v6).Style(myStyle2);
  ws3.Cell(34, 5, 35, 6, true).String(data1[6].v5).Style(myStyle3);
  ws3.Cell(34, 7, 35, 7, true).String(data1[6].v6).Style(myStyle2);
  ws3.Cell(35, 1, 35, 2, true).String(data1[4].v5).Style(myStyle3);
  ws3.Cell(35, 3, 35, 4, true).String(data1[4].v6).Style(myStyle2);
  ws3.Cell(36, 1, 36, 2, true).String(data1[5].v5).Style(myStyle3);
  ws3.Cell(36, 3, 36, 4, true).String(data1[5].v6).Style(myStyle2);
  ws3.Cell(36, 5, 37, 6, true).String(data1[3].v5).Style(myStyle3);
  ws3.Cell(36, 7, 37, 7, true).String(data1[3].v6).Style(myStyle2);
  ws3.Cell(37, 1, 37, 2, true).String(data1[7].v3).Style(myStyle3);
  ws3.Cell(37, 3, 37, 4, true).String(data1[7].v4).Style(myStyle2);

  ws3.Cell(38, 1, 38, 5, true).String('Заключение о возможности/невозможности эксплуатации транспортного средства').Style(myStyle4);
  ws3.Cell(39, 1, 39, 5, true).String('Results of the roadworthiness inspection').Format.Font.Family('Times New Roman');
  ws3.Cell(38, 6, 39, 6, true).String('Возможно   Passed').Style(myStyle);
  ws3.Cell(38, 7, 39, 7, true).String('Невозможно    Failed').Style(myStyle);
  if (post.result == '1')
    ws3.Cell(38, 6).Style(myStyle2);
  else if (post.result == '2')
    ws3.Cell(38, 7).Style(myStyle2);
  ws3.Cell(46, 3).String(post.date).Style(myStyle);
  ws3.Cell(40, 1, 42, 5, true).String('Пункты диагностической карты, требующие повторной проверки:').Format.Font.Alignment.Vertical('top').Format.Font.Family('Times New Roman');
  ws3.Cell(40, 6, 41, 7, true).String('Повторный технический контроль пройти до:').Style(myStyle);
  ws3.Cell(42, 6, 42, 7, true).String(post.repeat).Style(myStyle);
  ws3.Cell(44, 1, 45, 2, true).Format.Font.Alignment.Horizontal('center').String('Номер в ЕАИСТО');
  ws3.Cell(44, 3, 45, 4, true).Style(myStyle2);
  ws3.Cell(46, 1, 46, 2, true).Format.Font.Alignment.Horizontal('center').String('Дата проверки ТС:');
  ws3.Cell(46, 3, 46, 4, true).String(post.date).Style(myStyle2);
  ws3.Cell(46, 6).String('Печать         Stamp').Style(myStyle);
  ws3.Cell(47, 1, 47, 3, true).String('Ф.И.О. технического эксперта');
  ws3.Cell(47, 4).String(post.expert).Style(myStyle2);
  ws3.Cell(48, 1, 48, 3, true).String('Подпись                                   Signature').Style(myStyle);
  //Border
  ws3.Cell(1, 8, 49, 8).Style(leftBorder);
  ws3.Cell(49, 1, 49, 7).Style(bottomBorder);
}

function createFourthSheet(data1, post) {
  ws4 = wb.WorkSheet('квитанция', wsOpts);
  var myStyle = wb.Style();
  myStyle.Font.Size(12);
  myStyle.Font.Family('Times New Roman');
  myStyle.Font.Alignment.Vertical('center');
  myStyle.Font.Alignment.Horizontal('left');
  myStyle.Font.WrapText();
  var myStyle2 = wb.Style();
  myStyle2.Font.Size(12);
  myStyle2.Font.Family('Times New Roman');
  myStyle2.Font.Alignment.Vertical('center');
  myStyle2.Font.Alignment.Horizontal('center');
  myStyle2.Font.WrapText();
  var myBorder = wb.Style();
  myBorder.Border({
    top: {
      style: 'thick',
      color: '0000FF'
    },
    left: {
      style: 'thick',
      color: '0000FF'
    },
    right: {
      style: 'thick',
      color: '0000FF'
    },
    bottom: {
      style: 'thick',
      color: '0000FF'
    }
  });

  ws4.Column(2).Width(22);
  ws4.Column(7).Width(12);
  ws4.Column(8).Width(10);
  ws4.Column(9).Width(10);

  ws4.Cell(1, 1, 1, 7, true).String('Получатель: Индивидуальный предприниматель Пополитов Руслан Александрович');
  ws4.Cell(2, 1, 2, 7, true).String('Банк получателя: РОССИЙСКИЙ НАЦИОНАЛЬНЫЙ КОММЕРЧЕСКИЙ БАНК (ПАО)');
  ws4.Cell(3, 1, 3, 7, true).String('р/с: 40802810741200000048');
  ws4.Cell(4, 1, 4, 7, true).String('к/с: 30101810400000000607 в ОПЕРУ МГТУ Банка России');
  ws4.Cell(5, 1, 5, 7, true).String('БИК: 044525607  ИНН: 920100006420');
  ws4.Cell(6, 1, 6, 2, true).Format.Font.Bold().Format.Font.Alignment.Horizontal('center').String('Плательщик');
  ws4.Cell(6, 3, 6, 7, true).Format.Font.Bold().Format.Font.Alignment.Horizontal('center').String(data1[9].v2);
  ws4.Cell(7, 1, 7, 2, true).Format.Font.Alignment.Horizontal('center').String('моб. тел.');
  ws4.Cell(7, 3, 7, 7, true).Format.Font.Alignment.Horizontal('center').String('+79788987206');
  ws4.Cell(8, 1, 8, 2, true).Format.Font.Alignment.Horizontal('center').String('Вид платежа');
  ws4.Cell(8, 3, 8, 7, true).Format.Font.Alignment.Horizontal('center').String('за услуги по проведению технического  контроля');
  ws4.Cell(9, 1, 9, 2, true).Format.Font.Alignment.Horizontal('center').String('Cумма'); //wtf
  ws4.Cell(10, 1, 10, 2, true).Format.Font.Alignment.Horizontal('center').String(''); //wtf
  ws4.Cell(14, 1, 14, 9, true).String('Расписка в получении денежных средств').Style(myStyle2);
  ws4.Row(17).Height(40);
  ws4.Cell(16, 1, 17, 9, true).String('Я,  Соломатов Алексей Валерьевич, во исполнение поручения, совершенного в порядке предусмотренного главой 49 Гражданского кодекса РФ и установленном ст.16 закона Российской Федерации от 07.07.2011 г. №170-ФЗ  "О техническом контроле транспортных средств и о внесении изменений в отдельные законодательные акты РФ",')
    .Style(myStyle);
  ws4.Cell(18, 1).String('принял от').Style(myStyle);
  ws4.Cell(18, 2, 18, 3, true).String(data1[9].v2).Style(myStyle2);
  ws4.Cell(18, 4, 18, 5, true).String('денежные средства в размере').Style(myStyle); //wtf
  ws4.Cell(18, 8, 18, 9, true).String('для оплаты услуг по').Style(myStyle);
  ws4.Cell(19, 1, 19, 9, true).String(' техническому контролю. Обязуюсь внести их в полном объеме на расчетный счет ИП Пополитов Р.А. за проведение').Style(myStyle);
  ws4.Cell(20, 1, 20, 2, true).String(' технического контроля автомобиля ').Style(myStyle);
  ws4.Cell(20, 3, 20, 4, true).String(data1[4].v4).Style(myStyle2);
  ws4.Cell(20, 5, 20, 6, true).String('гос. регистрационный знак').Style(myStyle);
  ws4.Cell(20, 7).String(data1[4].v2).Style(myStyle2);
  ws4.Cell(21, 1, 21, 9, true).String(' в течение 3-х дней. С условиями приема указанной в платежном документе суммы, в т.ч. с суммой взымаемой платы ').Style(myStyle);
  ws4.Cell(22, 1, 22, 9, true).String('за услуги  банка, ознакомлен и согласен.').Style(myStyle);
  ws4.Cell(25, 2).String(post.date).Style(myStyle2);
  ws4.Cell(25, 6, 25, 9, true).String(' ________________ /А.В. Соломатов/').Style(myStyle);
  ws4.Cell(28, 1, 28, 9, true).String('_________________________________________________________________________________________________________________________').Style(myStyle);
  ws4.Cell(30, 1, 30, 9, true).String('Поручение на оплату услуг').Style(myStyle2);
  ws4.Cell(32, 1).String('Я,  ').Style(myStyle).Format.Font.Alignment.Horizontal('right');
  ws4.Cell(32, 2, 32, 3, true).String(data1[9].v2).Style(myStyle);
  ws4.Cell(32, 4, 32, 9, true).String('руководствуясь ст.16 закона Российской Федерации от 07.07.2011 г. №170-ФЗ ').Style(myStyle);
  ws4.Cell(33, 1, 33, 9, true).String('"О техническом осмотре транспортных средств и о внесении изменений в отдельные законодательные акты РФ",').Style(myStyle);
  ws4.Cell(34, 1, 34, 3, true).String('представляя для технического контроля автомобиль ').Style(myStyle);
  ws4.Cell(34, 4, 34, 5, true).String(data1[4].v4).Style(myStyle2);
  ws4.Cell(34, 6, 34, 7, true).String('гос. регистрационный знак').Style(myStyle2);
  ws4.Cell(34, 8, 34, 9, true).String(data1[4].v2).Style(myStyle);
  ws4.Cell(35, 1, 35, 9, true).String('в порядке, предусмотренном главой 49 Гражданского кодекса РФ поручаю внести за меня плату за оказания услуг').Style(myStyle);
  ws4.Cell(36, 1).String(' в размере').Style(myStyle); //wtf
  ws4.Cell(37, 1, 37, 9, true).String('С условиями приема указанной в платежном документе суммы, в т.ч. с суммой взымаемой платы за услуги банка, ознакомлен и согласен.').Style(myStyle);
  ws4.Cell(39, 2).String(post.date).Style(myStyle2);
  ws4.Cell(39, 5, 39, 9, true).String('____________________/' + data1[9].v2).Style(myStyle);
  //Border
  ws4.Cell(1, 1, 40, 9).Style(myBorder);
}

function final(post) {
  ws0.Cell(2, 27).String(post.reg);
  ws2.Cell(2, 5).String(post.reg);
  ws3.Cell(44, 3).String(post.reg);
}

var onRequest = function(req, res) {
  console.log("request received!");
  if (req.method == 'GET') {
    console.log('GET!!!!!');
    var url_parts = url.parse(req.url, true);
    var query = url_parts.query;
    console.log(query);
    if (query.path) {
      if (query.mypath) {
        path = query.mypath;
        fs.writeFile('path.txt', path, function(err) {
          if (err) {
            console.log('Ошибка! Закройте файл path.txt!');
            throw err;
          }
        });
      }
      res.end(path);
      return;
    } else if (query.journalPath) {
      if (query.journalPathNew) {
        journalPath = query.journalPathNew;
        fs.writeFile('journalPath.txt', journalPath, function(err) {
          if (err) {
            console.log('Ошибка! Закройте файл journalPath.txt!');
            throw err;
          }
        });
      }
      res.end(journalPath);
      return;
    }
  }
  if (req.method == 'POST') {
    var body = '';
    req.on('data', function(data) {
      body += data;
      // Too much POST data, kill the connection!
      if (body.length > 1e6)
        req.connection.destroy();
    });
    req.on('end', function() {
      var post = qs.parse(body);
      console.log(post);
      if (post.msg == 'step 1') {
        journal = wb = wsOpts = ws0 = ws = ws2 = ws3 = ws4 = '';
        data1 = createWorkBook(post);
        createFirstSheet();
        console.log('STEP 1');
      } else if (post.msg == 'step 2') {
        ws2 = '';
        diagnosticIssues = JSON.parse(post.diagnosticIssues);
        failedRequirements = JSON.parse(post.failedRequirements);
        notes = post.notes;
        createSecondSheet(data1, post);
        console.log('STEP 2');
      } else if (post.msg == 'step 3') {
        ws3 = '';
        createThirdSheet(data1, post);
        createFourthSheet(data1, post);
        console.log('STEP 3');
      } else {
        console.log('FINAL STEP');
        final(post);
        // Synchronously write file
        journal.write(path.normalize(journalPath));
        wb.write(path.normalize(wbPath + "/" + regNumber + ".xlsx"));
      }
      res.end();
    });
  }
};

http.createServer(onRequest).listen(serverPort, serverAddress);
console.log('Server running at ' + serverAddress + ":" + serverPort);
