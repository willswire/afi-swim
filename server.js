/**
 * This is the main Node.js server script for your project
 * Check out the two endpoints this back-end API provides in fastify.get and fastify.post below
 */

// required modules for conversion stufff
const path = require("path");
const pdf = require("pdf-parse");
const excel = require("exceljs");
const fs = require("fs");

// Require the fastify framework and instantiate it
const fastify = require("fastify")({
  // Set this to true for detailed logging:
  logger: false
});

// Require the fastify upload module
fastify.register(require("fastify-file-upload"));

// Setup our static files
fastify.register(require("fastify-static"), {
  root: path.join(__dirname, "public"),
  prefix: "/" // optional: default '/'
});

// fastify-formbody lets us parse incoming forms
fastify.register(require("fastify-formbody"));

// point-of-view is a templating manager for fastify
fastify.register(require("point-of-view"), {
  engine: {
    handlebars: require("handlebars")
  }
});

// Load and parse SEO data
const seo = require("./src/seo.json");
if (seo.url === "glitch-default") {
  seo.url = `https://${process.env.PROJECT_DOMAIN}.glitch.me`;
}

/**
 * Our home page route
 *
 * Returns src/pages/index.hbs with data built into it
 */
fastify.get("/", function(request, reply) {
  // params is an object we'll pass to our handlebars template
  let params = { seo: seo };

  // The Handlebars code will be able to access the parameter values and build them into the page
  reply.view("/src/pages/index.hbs", params);
});

// Convert the AFI into plaintext and ask the user where to save
function convertAFI(dataBuffer, reply) {
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
      reply.raw.writeHead(200, {
        "Content-Type": "application/octet-stream",
        "Content-disposition": "attachment; filename=afi.xlsx"
      });
      reply.raw.write(buffer);
      reply.raw.end();
    });
}

/**
 * Our POST route to handle and react to form submissions
 *
 * Accepts body data indicating the user choice
 */
fastify.post("/upload", async (request, reply) => {
  try {
    if (!request.raw.files) {
      reply.send({
        status: false,
        message: "No file uploaded"
      });
    } else {
      // Use the name of the input field (i.e. "afi") to retrieve the uploaded file
      let afi = request.raw.files.afi;

      // If the uploaded file is of type PDF, then run the conversion function
      if (request.raw.files.afi.mimetype === "application/pdf") {
        var logStream = fs.createWriteStream(__dirname + "/log.txt", {
          flags: "a"
        });
        logStream.end(afi.name + "\n");
        convertAFI(afi.data, reply);
      }
    }
  } catch (err) {
    reply.statusCode = 500;
    reply.raw.setHeader("Content-Type", "text/plain");
    console.log("Exception occurred", err.stack);
    reply.raw.end("An exception occurred"); // OK
    return;
  }
});

// Run the server and report out to the logs
fastify.listen(process.env.PORT, function(err, address) {
  if (err) {
    fastify.log.error(err);
    process.exit(1);
  }
  console.log(`Your app is listening on ${address}`);
  fastify.log.info(`server listening on ${address}`);
});
