var datass = '';
var DataArr = [];
pdfjsLib.GlobalWorkerOptions.workerSrc = './pdfjs/build/pdf.worker.js';

function ExtractText() {
  var input = document.getElementById("file-id");
  var fReader = new FileReader();
  fReader.readAsDataURL(input.files[0]);
  console.log(input.files[0]);
  fReader.onloadend = function (event) {
    convertDataURIToBinary(event.target.result);
  }
}

var BASE64_MARKER = ';base64,';

function convertDataURIToBinary(dataURI) {

  var base64Index = dataURI.indexOf(BASE64_MARKER) + BASE64_MARKER.length;
  var base64 = dataURI.substring(base64Index);
  var raw = window.atob(base64);
  var rawLength = raw.length;
  var array = new Uint8Array(new ArrayBuffer(rawLength));

  for (var i = 0; i < rawLength; i++) {
    array[i] = raw.charCodeAt(i);
  }
  pdfAsArray(array)

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

function pdfAsArray(array) {

  var loadingTask = pdfjsLib.getDocument({data: array});

  loadingTask.promise.then(function(pdf) {

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

      // Display text of all the pages in the console
      // e.g ["Text content page 1", "Text content page 2", "Text content page 3" ... ]
      console.log(pagesText); // representing every single page of PDF Document by array indexing
      console.log(pagesText.length);
      var outputStr = "";
      for (var pageNum = 0; pageNum < pagesText.length; pageNum++) {
        console.log(pagesText[pageNum]);
        outputStr = "";
        outputStr = "<br/><br/>Page " + (pageNum + 1) + " contents <br/> <br/>";

        var div = document.getElementById('output');

        div.innerHTML += (outputStr + pagesText[pageNum]);

      }

    });

  }, function (reason) {
    // PDF loading error
    console.error(reason);
  });
}