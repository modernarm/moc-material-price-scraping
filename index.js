import * as fs from "fs";
import { chromium } from "playwright";
import { Iconv } from "iconv";
import { load } from "cheerio";
import HtmlTableToJson from "html-table-to-json";
import _ from "lodash";
import xl from "excel4node";

const rowItem = {
  no: null,
  name: null,
  unit: null,
  jan: null,
  feb: null,
  mar: null,
  apr: null,
  may: null,
  jun: null,
  jul: null,
  aug: null,
  sept: null,
  oct: null,
  nov: null,
  dec: null,
  average: null,
};
let rawdata = fs.readFileSync("region.json");
let data = JSON.parse(rawdata);
const regions = ["1", "2", "3", "4"];
let materialRawData = fs.readFileSync("material.json");
let materials = [...JSON.parse(materialRawData)];
const iconv = new Iconv("CP874", "UTF-8");

const main = async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 300 });
  const page = await browser.newPage();
  for (let region of regions) {
    const uri =
      "http://www.indexpr.moc.go.th/PRICE_PRESENT/Select_regionCsi_xls.asp?region=" +
      region;
    await page.goto(uri);

    for (const province of data[region]) {
      await page.selectOption('select[name="DDProvince"]', province.code);
      const [download] = await Promise.all([
        page.waitForEvent("download"),
        page.click('input[type="submit"]'),
      ]);
      const path = await download.path();
      fs.copyFile(path, "./download/" + province.code + ".html", (err) => {
        if (err) throw err;
        console.log("copied to success");
      });
    }
  }
  page.close();
  readFile();
};

const convertEncodingFile = () => {
  for (let region of regions) {
    for (const province of data[region]) {
      fs.readFile(`./download/${province.code}.html`, function (err, data) {
        if (err) throw err;
        const buffer = iconv.convert(data);
        fs.writeFile(`./convert/${province.code}.html`, buffer, () => {});
      });
    }
  }
};

const tableToJson = () => {
  for (let region of regions) {
    for (const province of data[region]) {
      const $ = load(fs.readFileSync(`./convert/${province.code}.html`));
      const d2Items = $(".d2").contents();
      const provinceName = d2Items[0].data;
      const year = d2Items[1].data;
      const table = $("table");
      const jsonTables = HtmlTableToJson.parse(table.toString());
      let items = [];
      var wb = new xl.Workbook();
      var ws = wb.addWorksheet("Sheet 1");
      ws.cell(1, 1).string("no");
      ws.cell(1, 2).string("name");
      ws.cell(1, 3).string("unit");
      ws.cell(1, 4).string("jan");
      ws.cell(1, 5).string("feb");
      ws.cell(1, 6).string("mar");
      ws.cell(1, 7).string("apr");
      ws.cell(1, 8).string("may");
      ws.cell(1, 9).string("jun");
      ws.cell(1, 10).string("jul");
      ws.cell(1, 11).string("aug");
      ws.cell(1, 12).string("sept");
      ws.cell(1, 13).string("oct");
      ws.cell(1, 14).string("nov");
      ws.cell(1, 15).string("dec");
      ws.cell(1, 16).string("average");

      for (let o = 0; o < jsonTables.results[0].length; o++) {
        const current = { ...rowItem };
        const item = jsonTables.results[0][o];
        const row = o + 2;
        current.no = item[Object.keys(item)[0]];
        ws.cell(row, 1).string(current.no);
        current.name = String(item[Object.keys(item)[1]]).replace(
          / +(?= )/g,
          ""
        );
        ws.cell(row, 2).string(current.name);
        current.unit = item[Object.keys(item)[2]];
        ws.cell(row, 3).string(current.unit);
        current.jan = String(item[Object.keys(item)[3]]);
        ws.cell(row, 4).string(current.jan);
        current.feb = String(item[Object.keys(item)[4]]);
        ws.cell(row, 5).string(current.feb);
        current.mar = String(item[Object.keys(item)[5]]);
        ws.cell(row, 6).string(current.mar);
        current.apr = String(item[Object.keys(item)[6]]);
        ws.cell(row, 7).string(current.apr);
        current.may = String(item[Object.keys(item)[7]]);
        ws.cell(row, 8).string(current.may);
        current.jun = String(item[Object.keys(item)[8]]);
        ws.cell(row, 9).string(current.jun);
        current.jul = String(item[Object.keys(item)[9]]);
        ws.cell(row, 10).string(current.jul);
        current.aug = String(item[Object.keys(item)[10]]);
        ws.cell(row, 11).string(current.aug);
        current.sept = String(item[Object.keys(item)[11]]);
        ws.cell(row, 12).string(current.sept);
        current.oct = String(item[Object.keys(item)[12]]);
        ws.cell(row, 13).string(current.oct);
        current.nov = String(item[Object.keys(item)[13]]);
        ws.cell(row, 14).string(current.nov);
        current.dec = String(item[Object.keys(item)[14]]);
        ws.cell(row, 15).string(current.dec);
        current.average = String(item[Object.keys(item)[15]]);
        ws.cell(row, 16).string(current.average);
        items.push(current);
      }
      wb.write(`./excel/${province.code}-${provinceName}-${year}.xlsx`);
      fs.writeFile(
        `./json/province-${province.code}.json`,
        JSON.stringify(items),
        () => {}
      );
    }
  }
};
// main();
// convertEncodingFile();
tableToJson();
