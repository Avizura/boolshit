var serverAddress = 'http://127.0.0.1:5000';
var request = new XMLHttpRequest();

function toUrlEncoded(obj) {
  var urlEncoded = "";
  for (var key in obj) {
    urlEncoded += encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]) + '&';
  }
  return urlEncoded;
}

function getPath() {
  request.onreadystatechange = function() {
    if (request.readyState == 4 && request.status == 200) {
      document.getElementById('path').value = request.responseText;
    }
  }
  request.open('POST', serverAddress + '/mytest', true);
  request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  request.send(toUrlEncoded({
    path: true,
    mypath: document.getElementById('path').value
  }));
}
getPath();

function go() {
  document.getElementById('go').innerHTML = 'OK!';
  chrome.tabs.query({
    active: true,
    currentWindow: true
  }, function(tabs) {
    chrome.tabs.sendMessage(tabs[0].id, {
      tab: tabs[0]
    }, function(response) {
      document.getElementById("mytext").innerHTML = "SUCCESS!";
      alert("Success");
    });
  });
}

document.getElementById('go').addEventListener('click', go);
document.getElementById('path').addEventListener('blur', getPath);
