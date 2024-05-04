const { chromium } = require("playwright");
const { createObjectCsvWriter } = require("csv-writer");
const { Workbook } = require("exceljs");
const readline = require("readline");
const fs = require("fs");
const csv = require("fast-csv");
const ProgressBar = require("progress");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

async function excelPrompt() {
  return new Promise((resolve, reject) => {
    rl.question(
      "Would you like a summary of each article? (Y/N): ",
      (answer) => {
        const trimmedAnswer = answer.trim().toLowerCase();
        if (trimmedAnswer === "y") {
          resolve(true);
        } else if (trimmedAnswer === "n") {
          resolve(false);
        } else {
          console.log("Invalid input. Please type 'Y' for yes or 'N' for no.");
          excelPrompt().then(resolve);
        }
      }
    );
  });
}

async function scrapeHackerNews() {
  // Launch Browser
  const browser = await chromium.launch({ headless: false });
  const page = await browser.newPage();

  // Navigate to Hacker News
  await page.goto("https://news.ycombinator.com");

  // Extract top 10 articles
  const articleElements = await page.$$(".athing");

  // Scraping articles in parallel
  const articlePromises = articleElements.slice(0, 10).map(async (element) => {
    return await page.evaluate((elem) => {
      const rank = elem.querySelector(".rank").textContent;
      const title = elem
        .querySelector(".titleline a")
        .textContent.replace(";", ","); // prevents issues if ";" is inside the title
      const url = elem.querySelector(".titleline a").href;
      return { rank, title, url };
    }, element);
  });
  const articles = await Promise.all(articlePromises);

  // Write the articles to a CSV file
  const csvWriter = createObjectCsvWriter({
    path: "hacker-news-top-10-articles.csv",
    header: [
      { id: "rank", title: "Rank" },
      { id: "title", title: "Article Title" },
      { id: "url", title: "URL" },
    ],
    fieldDelimiter: ";", // Comma delimiter
  });
  await csvWriter.writeRecords(articles);

  // Close Browser
  await browser.close();

  return articles; // Return the articles array
}

async function createExcelFile() {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet("Hacker News Top 10 Articles");

  // Read data from CSV using fast-csv
  const records = [];
  fs.createReadStream("hacker-news-top-10-articles.csv")
    .pipe(csv.parse({ delimiter: ";" })) // Specify the delimiter
    .on("data", (data) => {
      records.push(data);
    })
    .on("end", () => {
      // Add data to Excel worksheet
      records.forEach((record) => {
        worksheet.addRow(record);
      });

      // Write Excel file
      workbook.xlsx.writeFile("hacker-news-top-10-articles.xlsx");
    });
  return { workbook, worksheet }; // Return both workbook and worksheet object
}

async function updateExcelWithSummaries(workbook, worksheet, articles) {
  const browser = await chromium.launch({ headless: false });
  const page = await browser.newPage();

  const progressBar = new ProgressBar(":bar :percent", {
    total: articles.length,
    width: 30,
  });

  console.log("AI Summaries powered by SummarAIse\u2122!");
  console.log("For more information, visit: https://summaraise.netlify.app/");

  for (let i = 0; i < articles.length; i++) {
    const article = articles[i];
    const url = article.url;

    await page.goto("https://summaraise.netlify.app/");
    await page.fill(".url_input", url);
    await page.click(".submit_btn");

    try {
      await page.waitForSelector(".summary_box", {
        visible: true,
        timeout: 10000,
      });
      let summary = await page.$eval(".summary_box", (element) =>
        element.textContent.trim()
      );

      if (summary.startsWith("Something wrong happened..")) {
        summary = "Summary Unavailable.";
      } else {
        // Truncate summary to 100 words and ensure it ends with a full stop
        const summaryWords = summary.split(" ");
        let truncatedSummary = summaryWords.slice(0, 400).join(" ");
        // Find the index of the last occurrence of a full stop within the first 400 words
        const lastFullStopIndex = truncatedSummary.lastIndexOf(".");
        if (lastFullStopIndex !== -1) {
          truncatedSummary = truncatedSummary.substring(
            0,
            lastFullStopIndex + 1
          ); // Include the last full stop
        }
        summary = truncatedSummary;
      }

      worksheet.getRow(i + 2).getCell("D").value = summary;
    } catch (error) {
      worksheet.getRow(i + 2).getCell("D").value =
        "Summary Unavailable, URL page could not be accessed.";
    }

    // Update progress bar
    progressBar.tick();
  }

  // Ensures URL column is clickable with common display text "LINK" for all rows except Column Name
  worksheet
    .getColumn("C")
    .eachCell({ includeEmpty: true }, (cell, rowNumber) => {
      if (rowNumber > 1 && cell.value && typeof cell.value === "string") {
        const url = cell.value;
        const hyperlinkFormula = `=HYPERLINK("${url}", "LINK")`;
        const hyperlink = worksheet.getCell(`C${rowNumber}`);
        hyperlink.value = { formula: hyperlinkFormula };
        hyperlink.font = { color: { argb: "0000FF" }, underline: true };
      }
    });

  // Set column headers
  worksheet.getRow(1).values = ["Rank", "Article Title", "URL", "AI Summary"];

  // Dynamic column widths
  worksheet.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const cellLength = cell.value ? cell.value.toString().length : 10;
      maxLength = Math.max(maxLength, cellLength);
    });
    column.width = Math.max(10, maxLength - 3.55);

    // Rank Row Bold
    worksheet.getRow(1).font = { bold: true };
    // "AI Summary" Column Formatting
    worksheet.getColumn(4).width = 160;

    // Loop through rows A2 to A11 and set Row Height
    for (let i = 2; i <= 11; i++) {
      const row = worksheet.getRow(i);
      row.height = 165;
    }

    const columnFields = ["A", "B", "C", "D"];
    const alignments = [
      { horizontal: "center", vertical: "middle" },
      { horizontal: "center", vertical: "middle" },
      { horizontal: "center", vertical: "middle" },
      { horizontal: "left", vertical: "middle", wrapText: true },
    ];
    columnFields.forEach((header, index) => {
      const column = worksheet.getColumn(header);
      column.alignment = alignments[index];
    });
    // Header Formatting
    const columnHeaders = ["A", "B", "C", "D"];

    columnHeaders.forEach((header) => {
      const columnName = worksheet.getCell(`${header}1`);
      columnName.alignment = { horizontal: "center", vertical: "middle" };
      columnName.font = { bold: true };
    });
  });
  // Excel formatting End

  // Write Excel file
  workbook.xlsx
    .writeFile("hacker-news-top-10-articles.xlsx")
    .then(() => {
      console.log("Excel file created successfully!");
      // Delete CSV file
      fs.unlinkSync("hacker-news-top-10-articles.csv");
    })
    .catch((err) => {
      console.error("Error writing Excel file:", err);
    });
  await browser.close();
}

(async () => {
  try {
    const buildExcel = await excelPrompt();
    if (buildExcel) {
      const articles = await scrapeHackerNews(); // Await scrapeHackerNews function
      const { workbook, worksheet } = await createExcelFile(); // Destructure workbook and worksheet
      await updateExcelWithSummaries(workbook, worksheet, articles); // Pass the worksheet object to the function
      console.log("Excel file and AI Summaries created successfully!");
    } else {
      await scrapeHackerNews();
      console.log("CSV file created successfully!");
    }
  } catch (error) {
    console.error("Error:", error);
  } finally {
    rl.close();
  }
})();
