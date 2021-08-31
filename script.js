pdfjsLib.GlobalWorkerOptions.workerSrc = './pdfjs/pdf.worker.js';

function swim() {
  var input = document.getElementById("file-id");
  var fileReader = new FileReader();
  fileReader.readAsDataURL(input.files[0]);
  fileReader.onloadend = function (event) {
    convertToBinary(event.target.result);
  }
}

function convertToBinary(dataURI) {
  const BASE64_MARKER = ';base64,';
  var base64Index = dataURI.indexOf(BASE64_MARKER) + BASE64_MARKER.length;
  var base64 = dataURI.substring(base64Index);
  var pdfData = atob(base64);
  pdfAsArray(pdfData)
}

function getPageText(pageNum, PDFDocumentInstance) {
  // Return a Promise that is solved once the text of the page is retrieven
  return new Promise(function (resolve, reject) {
    PDFDocumentInstance.getPage(pageNum).then(function (pdfPage) {
      // The main trick to obtain the text of the PDF page, use the getTextContent method
      pdfPage.getTextContent().then(function (textContent) {
        var textItems = textContent.items;
        var finalString = "";

        // Concatenate the string of the item to the final string
        for (var i = 0; i < textItems.length; i++) {
          var item = textItems[i];

          finalString += item.str + " ";
        }

        // Solve promise with the text retrieven from the page
        resolve(finalString);
      });
    });
  });
}

function pdfAsArray(pdfData) {

  var loadingTask = pdfjsLib.getDocument({ data: pdfData });

  loadingTask.promise.then(function (pdf) {

    var pdfDocument = pdf;
    // Create an array that will contain our promises
    var pagesPromises = [];

    for (var i = 0; i < pdf.numPages; i++) {
      // Required to prevent that i is always the total of pages
      (function (pageNumber) {
        // Store the promise of getPageText that returns the text of a page
        pagesPromises.push(getPageText(pageNumber, pdfDocument));
      })(i + 1);
    }

    // Execute all the promises
    Promise.all(pagesPromises).then(function (pagesText) {

      var raw = "";
      for (var pageNum = 0; pageNum < pagesText.length; pageNum++) {
        raw += pagesText[pageNum]
      }

      var regex = /\s(\d+\.)+\s/g;
      var output = raw.replace(regex, `<br><br>$&`)

      var div = document.getElementById('output');
      div.innerHTML = output;
    });

  }, function (reason) {
    // PDF loading error
    console.error(reason);
  });
}

// Excel Testing

var wb = XLSX.utils.book_new();
wb.Props = {
  Title: "Test Workbook",
  Subject: "Test",
  Author: "AFI SWIM",
  CreatedDate: new Date()
};

wb.SheetNames.push("Test Sheet");

var ws_data = [['hello', 'world']];  //a row with 2 columns

var ws = XLSX.utils.aoa_to_sheet(ws_data);

wb.Sheets["Test Sheet"] = ws;

var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

function s2ab(s) {
  var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
  var view = new Uint8Array(buf);  //create uint8array as viewer
  for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
  return buf;
}

function download() {
  saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), 'test.xlsx');
}