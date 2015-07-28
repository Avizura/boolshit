var serverAddress = 'http://127.0.0.1:5000';

function toUrlEncoded(obj) {
  var urlEncoded = "";
  for (var key in obj) {
    urlEncoded += encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]) + '&';
  }
  return urlEncoded;
}

function getPath() {
  var request = new XMLHttpRequest();
  request.onreadystatechange = function() {
    if (request.readyState == 4 && request.status == 200) {
      document.getElementById('path').value = request.responseText;
    }
  }
  request.open('GET', serverAddress + "?" + toUrlEncoded({path: true, mypath: document.getElementById('path').value}), true);
  request.send();
  // request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  // request.send(toUrlEncoded({
  //   path: true,
  //   mypath: document.getElementById('path').value
  // }));
}
getPath();

function getJournalPath(){
  var request = new XMLHttpRequest();
  request.onreadystatechange = function() {
    if (request.readyState == 4 && request.status == 200) {
      document.getElementById('journalPath').value = request.responseText;
    }
  }
  request.open('GET', serverAddress + "?" + toUrlEncoded({journalPath: true, journalPathNew: document.getElementById('journalPath').value}), true);
  request.send();
}
getJournalPath();

function go() {
  chrome.tabs.query({
    active: true,
    currentWindow: true
  }, function(tabs) {
    chrome.tabs.sendMessage(tabs[0].id, {
      tab: tabs[0]
    }, function(response) {});
  });
  this.innerHTML = "OK!";
}

document.getElementById('go').addEventListener('click', go);
document.getElementById('path').addEventListener('blur', getPath);
document.getElementById('journalPath').addEventListener('blur', getJournalPath);
