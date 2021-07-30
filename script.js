function save(element, filename = "") {
  var header =
    "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'><head><meta charset='UTF-8'><title>Document</title></head><body>";

  var footer = "</body></html>";

  var html = header + document.getElementById(element).innerHTML + footer;

  var blob = new Blob(["\ufeff", html], {
    type: "application/msword",
  });

  // Specify link url
  var url = "data:application/vnd.ms-word;charset=utf-8," + encodeURIComponent(html);

  // Specify file name
  filename = filename ? filename + ".docx" : "document.docx";

  // Create download link element
  var downloadLink = document.createElement("a");

  document.body.appendChild(downloadLink);

  if (navigator.msSaveOrOpenBlob) {
    navigator.msSaveOrOpenBlob(blob, filename);
  } else {
    // Create a link to the file
    downloadLink.href = url;

    // Setting the file name
    downloadLink.download = filename;

    //triggering the function
    downloadLink.click();
  }

  document.body.removeChild(downloadLink);
}

function execCmd(command) {
  document.execCommand(command, false, null);
}

function execCommandWithArg(command, arg) {
  document.execCommand(command, false, arg);
}

function printdiv(printpage) {
  var headstr = `<!DOCTYPE html>
<html>
  <head>
    <title></title>
  </head>
  <body>
  `;
  var footstr = `</body>
</html>`;
  var newstr = document.getElementById(printpage).innerHTML;
  var oldstr = document.body.innerHTML;
  document.body.innerHTML = headstr + newstr + footstr;
  window.print();
  document.body.innerHTML = oldstr;
  return false;
}
