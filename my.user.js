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
  return document.getElementById(id).value;
}

function geName(name) {
  return document.getElementsByName(name)[0].value;
}

function getSelectedValue(id) {
  var element = document.getElementById(id);
  return element.options[element.selectedIndex].innerHTML;
}

function getStep2Value(id) {
  var el = document.getElementById('parameter_row_' + id);
  if (el.classList.contains('tr_sel_disabled') || el.classList.contains('nonclickable'))
    return ' ';
  if (el.classList.contains('tr_sel_red'))
    return '-';
  return '+';
}

function getRadioValue() {
  if (document.getElementById('radio_0').checked == true)
    return '1';
  // else if (document.getElementById('radio_1').checked == true)
  return '2';
}

function getResult() {
  if (document.getElementsByClassName('zakluchenie_variant zakluchenie_da')[0].style.display == "")
    return '1';
  return '2';
}

function diagnisticResults() {
  var diagnosticIssues = [];
  var array = document.querySelectorAll('tr[id^="diagnostika_row_"]');
  for (var i = 0; i < array.length; i++) {
    var temp = array[i].children;
    var item = {
      v1: temp[0].firstChild.value,
      v2: temp[1].firstChild.value,
      v3: temp[2].firstChild.value,
      v4: temp[3].firstChild.value,
      v5: temp[4].children[0].options[temp[4].children[0].selectedIndex].innerHTML
    };
    diagnosticIssues.push(item);
  }
  return diagnosticIssues;
}

function failedRequirements() {
  var failedRequirements = [];
  var array = document.querySelectorAll('tr[id^="nevypoln_row_"]');
  for (var i = 0; i < array.length; i++) {
    var temp = array[i].children;
    var item = {
      v1: temp[0].firstChild.value,
      v2: temp[1].firstChild.value,
      v3: temp[2].children[0].options[temp[2].children[0].selectedIndex].innerHTML
    };
    failedRequirements.push(item);
  }
  return failedRequirements;
}

chrome.runtime.onMessage.addListener(function(msg, sender, sendResponse) {
  var validity = repeat = '';
  try {
    var step = document.getElementsByClassName("sel")[1].getElementsByClassName('kvadr')[0].innerHTML;
  } catch (e) {
    console.log(e);
  }
  if (step == '1') {
    var data = {
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
      tire: ge('MARKA_SHIN'),
      maxWeight: geName('RAZRESH_MAKS_MASSA'),
      fuel: getSelectedValue('TIP_TOPLIVA'),
      brakeSystem: getSelectedValue('TIP_TORMOZ_SISTEMY'),
      regSERIA: ge('SVID_O_REG_SERIA'),
      regNOMER: ge('SVID_O_REG_NOMER'),
      regKOGDA: ge('SVID_O_REG_KOGDA'),
      regKEM: geName('SVID_O_REG_KEM')
    };
  } else if (step == '2') {
    var data = {
        msg: "step 2",
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
        _65: getStep2Value(65),
        diagnosticIssues: JSON.stringify(diagnisticResults()),
        failedRequirements: JSON.stringify(failedRequirements()),
        notes: geName('RESULTAT_PRIMECHANIE')
      }
  } else if (step == '3') {
    if (getResult() == true)
      validity = ge('SROK_DEISTV');
    else repeat = geName('PROITI_POVTORNO_DO');
    data = {
      msg: 'step 3',
      result: getResult(),
      validity: validity,
      repeat: repeat,
      notes: ge('OSOB_OTMETKI'),
      checkType: getRadioValue(),
      date: document.getElementsByClassName('gray_tbl')[6].getElementsByTagName('td')[1].innerHTML,
      expert: document.getElementsByClassName('top_block_2')[0].getElementsByTagName('strong')[0].innerHTML
    };
  } else {
    data = {
      reg: document.getElementsByClassName('second_cont')[0].getElementsByTagName('h2')[1].getElementsByTagName('strong')[0].innerHTML
    };
  }
  request.onreadystatechange = function() {
    if (request.readyState == 4 && request.status == 200) {
      sendResponse();
    }
  }
  request.open('POST', serverAddress + '/mytest', true);
  request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  request.send(toUrlEncoded(data));
});
