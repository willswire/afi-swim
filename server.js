const express = require("express");
const fileUpload = require("express-fileupload");
const cors = require("cors");
const bodyParser = require("body-parser");
const morgan = require("morgan");
const _ = require("lodash");
const pdf = require("pdf-parse");
const excel = require("exceljs");
const app = express();
const fs = require("fs");
var RateLimit = require("express-rate-limit");

// set up limiter to prevent DDoS
var limiter = new RateLimit({
  windowMs: 1 * 60 * 1000, // 1 minute
  max: 5
});

// enable file upload
app.use(
  fileUpload({
    createParentPath: true,
    limits: {
      fileSize: 25 * 1024 * 1024 * 1024 // 25MB max file(s) size
    }
  })
);

//add other middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(morgan("dev"));

// apply rate limiter to all requests
app.use(limiter);

app.get("/", (request, response) => {
  response.sendFile(__dirname + "/index.html");
});

// Convert the AFI into plaintext and ask the user where to save
function convertAFI(dataBuffer, res) {
  //console.log("Converting AFI to Excel...")
  pdf(dataBuffer)
    .then(data => {
      // clean the text by removing all line breaks before paragraph numberings
      var cleanedText = data.text.replace(/\n[^\d+\.]/g, "");

      // deliminate the text by newline breaks
      var segmentedText = cleanedText.split(/\n/);

      // find the first paragraph of the AFI
      var found = false;
      var index = 0;
      while (!found) {
        if (segmentedText[index].includes("1.1.  ")) {
          found = true;
        } else {
          index++;
        }
      }

      // remove everything before paragraph 1.1. of the AFI
      segmentedText.splice(0, index);

      // pass off the result to the exceljs handler
      return segmentedText;
    })
    .then(result => {
      // Create our worksheet and set it up
      const workbook = new excel.Workbook();
      const worksheet = workbook.addWorksheet("AFI Export");
      worksheet.columns = [
        { header: "Description", key: "description", width: 70 },
        { header: "Shall/Will/Must", key: "shallwillmust", width: 30 },
        { header: "Tier Level", key: "tierlevel", width: 30 }
      ];

      // Set up the description column from the PDF export
      worksheet.getColumn(1).values = result;
      worksheet.getColumn(1).alignment = { wrapText: true };

      // Set up the SWM column by iterating over the rows, and then create the filter
      const shallwillmust = worksheet.getColumn(2);
      shallwillmust.eachCell(function(cell, rowNumber) {
        cell.value = {
          formula:
            '=CONCATENATE(IF(IFERROR(FIND(" shall ",A' +
            rowNumber +
            '),0)>0,"Shall",""),IF(IFERROR(FIND(" will ",A' +
            rowNumber +
            '),0)>0,"Will",""),IF(IFERROR(FIND(" must ",A' +
            rowNumber +
            '),0)>0,"Must",""))'
        };
      });

      // Set up the Tier column by iterating over the rows, and then create the filter
      const tier = worksheet.getColumn(3);
      tier.eachCell(function(cell, rowNumber) {
        cell.value = {
          formula:
            '=CONCATENATE(IF(IFERROR(FIND("T-0",A' +
            rowNumber +
            '),0)>0,"T-0",""),IF(IFERROR(FIND("T-1",A' +
            rowNumber +
            '),0)>0,"T-1",""),IF(IFERROR(FIND("T-2",A' +
            rowNumber +
            '),0)>0,"T-2",""),IF(IFERROR(FIND("T-3",A' +
            rowNumber +
            '),0)>0,"T-3",""))'
        };
      });

      // Insert the header row,, style, then freeze it
      worksheet.getColumn("description").header = "Description";
      worksheet.getColumn("shallwillmust").header = "Shall/Will/Must";
      worksheet.getColumn("tierlevel").header = "Tier Level";
      worksheet.getRow(1).font = { bold: true };
      worksheet.views = [{ state: "frozen", xSplit: 0, ySplit: 1 }];
      worksheet.autoFilter = "B1:C1";

      // return the workbook
      return workbook;
    })
    .then(workbook => {
      // write the workbook to the buffer
      return workbook.xlsx.writeBuffer();
    })
    .then(buffer => {
      // send a response to the original POST HTTP request
      // of the converted PDF now stored in the buffer
      res.writeHead(200, {
        "Content-Type": "application/octet-stream",
        "Content-disposition": "attachment; filename=afi.xlsx"
      });
      res.write(buffer);
      res.end();
    });
}

// This route is triggered when the form is submitted from the main page.
// The visitor will see the browser loading as the POST response is dependent
// on the file returned from the backend server
app.post("/upload", async (req, res) => {
  try {
    if (!req.files) {
      res.send({
        status: false,
        message: "No file uploaded"
      });
    } else {
      // Use the name of the input field (i.e. "afi") to retrieve the uploaded file
      let afi = req.files.afi;

      // If the uploaded file is of type PDF, then run the conversion function
      if (req.files.afi.mimetype === "application/pdf") {
        var logStream = fs.createWriteStream(__dirname + "/log.txt", {
          flags: "a"
        });
        logStream.end(afi.name + "\n");
        convertAFI(afi.data, res);
      }
    }
  } catch (err) {
    res.statusCode = 500;
    res.setHeader("Content-Type", "text/plain");
    console.log("Exception occurred", err.stack);
    res.end("An exception occurred"); // OK
    return;
  }
});

// Run the application on port 3000
app.listen(3000);
