import puppeteer from "puppeteer-extra";
import StealthPlugin from "puppeteer-extra-plugin-stealth";
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import os from "os";

puppeteer.use(StealthPlugin());

puppeteer.use(
  StealthPlugin({
    enabledEvasions: new Set([
      "chrome.app",
      "chrome.csi",
      "defaultArgs",
      "navigator.plugins",
    ]),
  })
);

const westernUnionFcaPath =
  "https://register.fca.org.uk/s/firm?id=0010X00004EMHyEQAX";

// Clickable button selectors
const allowAllCookiesSelector =
  "#modal-content-id-1 > footer > div > button:nth-child(3)";
const activitiesAndServicesSelector =
  "#who-is-this-firm-connected-to-appointed-representatives-button";
const showAllFirmsInListSelector =
  "#appointed-rep-table-resultcountselect-button";
const resultCountSelector = "#appointed-rep-table-result-count";

// Firm detail selectors
const firmNameSelector = "head > title";
const addressSelector =
  "#who-is-this-details-content > div.stack.stack--direct.stack--medium > div.slds-grid.slds-wrap.gutters-large.gutters-medium_none > div.slds-col.slds-size_1-of-1.slds-medium-size_6-of-12.slds-p-around_none.slds-p-right_small > div > div > div:nth-child(1) > p";
const zipCodeSelector =
  "#who-is-this-details-content > div.stack.stack--direct.stack--medium > div.slds-grid.slds-wrap.gutters-large.gutters-medium_none > div.slds-col.slds-size_1-of-1.slds-medium-size_6-of-12.slds-p-around_none.slds-p-right_small > div > div > div:nth-child(1) > p > span:nth-child(4)";
const phoneSelector =
  "#who-is-this-details-content > div.stack.stack--direct.stack--medium > div.slds-grid.slds-wrap.gutters-large.gutters-medium_none > div.slds-col.slds-size_1-of-1.slds-medium-size_6-of-12.slds-p-around_none.slds-p-right_small > div > div > div:nth-child(2) > p";
const emailSelector =
  "#who-is-this-details-content > div.stack.stack--direct.stack--medium > div.slds-grid.slds-wrap.gutters-large.gutters-medium_none > div.slds-col.slds-size_1-of-1.slds-medium-size_6-of-12.slds-p-around_none.slds-p-right_small > div > div > div:nth-child(3) > p";
const firmRefNumberSelector =
  "#who-is-this-details-content > div.stack.stack--direct.stack--medium > div.slds-grid.slds-wrap.gutters-large.gutters-medium_none > div:nth-child(3) > div > div > div:nth-child(1) > p";
const typeSelector = 
  "#who-is-this-status-content > div > div:nth-child(1) > div > div > div > p";
const agentStatusSelector =
  "#who-is-this-status-content > div > div:nth-child(2) > div > div > div > p:nth-child(2)";
  const dateSinceStatusSelector =
    "#who-is-this-status-content > div > div:nth-child(2) > div > div > div > p:nth-child(3)";

// Occasionally occurs in firm details
const registeredCompanyNumberSelector =
  "#who-is-this-details-content > div.stack.stack--direct.stack--medium > div.slds-grid.slds-wrap.gutters-large.gutters-medium_none > div:nth-child(3) > div > div > div:nth-child(2) > a";

let hrefArray = [];

export default async function scrapeDetailsAboutAllFirms() {
  const browser = await puppeteer.launch({
    headless: true,
    setTimeout: 120000,
  });
  const page = await browser.newPage();

  await page.goto(westernUnionFcaPath, {
    waitUntil: "networkidle0",
  });

  if ((await page.$(allowAllCookiesSelector)) !== null) {
    await page.click(allowAllCookiesSelector);
    console.log(
      ". . : : Scraper takes approximately 45 minutes to run : : . .\n"
    );
    console.log("Allowing cookies");
  }

  await openLargeListOfFirms();

  await saveAllFirmDetailLinksToArray();

  await traverseLinksAndScrapeDetails();

  async function openLargeListOfFirms() {
    await page.locator(activitiesAndServicesSelector).click();
    console.log("Clicking 'Appointed representatives and agents' block");

    console.log(
      "Clicking 'Show all results' to bring up large list of firms\nLoading for a few seconds ...\n"
    );
    //await page.locator(showAllFirmsInListSelector).click();
  }

  async function saveAllFirmDetailLinksToArray() {
    const resultText = await page.$eval(
      resultCountSelector,
      (el) => el.textContent
    );
    const totalAmountOfFirmsInList = parseInt(
      resultText.match(/out of (\d+)/)[1],
      10
    );
    console.log("Total number of firms in the list:", totalAmountOfFirmsInList);
    console.log("Grabbing " + totalAmountOfFirmsInList + " links\n");

    for (let i = 1; i < 5; i++) {
      // // // - - - Amount of firms to scrape
      const firmLinkSelector = `#appointed-rep-table-pagination-captured-table-${i}-cell-0-link`;
      try {
        const firmLinkElement = await page.$(firmLinkSelector);

        if (firmLinkElement) {
          const href = await page.evaluate(
            (element) => element.href,
            firmLinkElement
          );
          hrefArray.push(href);
        } else {
          console.log(`Firm link at index ${i} not found.`);
        }
      } catch (error) {
        console.error(`Error processing firm link at index ${i}:`, error);
      }
    }
  }

  async function traverseLinksAndScrapeDetails() {
          //                   EXCEL AND FILE PATHING - - - - - - - - - - - - - - -
    const today = new Date();
    const dateStr = today.toISOString().split("T")[0];
    const desktopPath = path.join(os.homedir(), "Desktop");
    const folderPath = path.join(desktopPath, "FCA scraper");

    if (!fs.existsSync(folderPath)) {
      fs.mkdirSync(folderPath);
    }

    const filePath = path.join(folderPath, `firms_${dateStr}.xlsx`);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Firm Details");

    worksheet.columns = [
      { header: "Firm Name", key: "firmName", width: 26 },
      { header: "Address", key: "address", width: 72 },
      { header: "ZIP code", key: "zipCode", width: 15 },
      { header: "Phone", key: "phone", width: 20 },
      { header: "Email", key: "email", width: 30 },
      { header: "Ref #", key: "refNumber", width: 12 },
      {
        header: "Registered company number",
        key: "registeredCompanyNumber",
        width: 34,
      },
      { header: "Type", key: "type", width: 10 },
      { header: "Agent Status", key: "agentStatus", width: 27 },
    ];

    worksheet.getRow(1).font = {
      name: "Arial", 
      bold: true, 
      size: 13, 
      underline: true, 
    };

    worksheet.views = [
      {
        state: "frozen",
        xSplit: 0,
        ySplit: 1,
        topLeftCell: "A2",
        activeCell: "A2",
      },
    ];
// --------------------------------------------------------------------------
    console.log(". . : : Scraping start : : . . \n");

    for (let i = 0; i < hrefArray.length; i++) {
      const href = hrefArray[i];

      try {
        // Check valid href
        if (!href || !href.startsWith("http")) {
          console.error(`Invalid href: ${href}`);
          continue;
        }

        console.log(`${i}: Navigating to ${href}`);
        await page.goto(href, { waitUntil: "networkidle0" });

        // Scrape data
        const firmName = await page.evaluate((selector) => { // Not super happy with violating DRY but works
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "-";
        }, firmNameSelector);

        const addressHtml = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.innerHTML : "-";
        }, addressSelector);

        const parsedAddress = parseAddressHtml(addressHtml);

        const zipCode = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "-";
        }, zipCodeSelector);

        const phone = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "-";
        }, phoneSelector);

        const email = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "-";
        }, emailSelector);

        const firmRefNumber = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "-";
        }, firmRefNumberSelector);

        const type = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "-";
        }, typeSelector);

        const agentStatus = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "-";
        }, agentStatusSelector);

        const dateSinceStatus = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : ""
        }, dateSinceStatusSelector);
        
        const registeredCompanyNumber = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          if (element) {
            const textContent = element.textContent.trim();
            return textContent.replace(/\D/g, ""); // Remove non-digits
          } else {
            return "-";
          }
        }, registeredCompanyNumberSelector);

        worksheet.addRow([
          firmName,
          parsedAddress,
          zipCode,
          phone,
          email,
          firmRefNumber,
          registeredCompanyNumber,
          type,
          agentStatus + " " + dateSinceStatus,
        ]);
      } catch (error) {
        console.error(`Error processing ${href}:`, error);
      }
    }

    await workbook.xlsx.writeFile(filePath);
    console.log(
      `\n . . : : Done scraping : : . .\n \n Excel file saved to: ${filePath}\n`
    );

    await browser.close();
  }

}

function parseAddressHtml(inputString) {
  let text = "";
  let closingTagCount = 0;
  let tagsToAvoid = [2, 3, 4, 5, 6];
  let textChunks = [];

  for (let i = inputString.length; i >= 0; i--) {
    if (inputString[i] == ">") {
      closingTagCount++;

      if (tagsToAvoid.includes(closingTagCount)) {
        // ???
        continue;
      }

      let chunk = "";
      let j = i + 1;

      while (j < inputString.length && inputString[j] !== "<") {
        chunk += inputString[j];
        j++;
      }

      textChunks.push(chunk.trim());
    }
  }

  let lastChunk = "";
  let k = 0;

  while (inputString[k] !== "<") {
    lastChunk += inputString[k];
    k++;
  }

  textChunks.push(lastChunk);

  let parseAddress = textChunks.join(", ");

  return parseAddress;
}

function delay(time) {
  return new Promise(function (resolve) {
    setTimeout(resolve, time);
  });
}
