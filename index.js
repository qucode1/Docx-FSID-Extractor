(async () => {
  const mammoth = require("mammoth");
  const x1 = require("excel4node");
  const fs = require("fs");

  // const idRe = /(?<=<h3>.*<table><tr><td><p>FS ID.*<tr><td><p>)\d*(?=<\/p><\/td>)/gm;
  // const withTableRe = /(?<=<h(3|2)>(<.*>)?)[^<>]+(?=<\/h(3|2)><table><tr><td><p>FS ID)/gm;
  // const combinedRe = /((?<=<h(3|2)>(<.*>)?)[^<>]+(?=<\/h(3|2)><table><tr><td><p>FS ID)|(?<=<h3>.*<table><tr><td><p>FS ID.*<tr><td><p>)\d*(?=<\/p><\/td>))/gm;
  // const headlineRe = /(?<=<h(1|2|3)>(<.*>)?)[^<>]+(?=<\/h(1|2|3)>)/gm;
  // const headlineIDRe = /((?<=<h(1|2|3)>(<.*>)?)[^<>]+(?=<\/h(1|2|3)>)|((?<=<h3>.*<table><tr><td><p>FS ID.*<tr><td><p>)\d*(?=<\/p><\/td>)))/gm;
  const h1Re = /(?<=<h1>(<.*>)?)[^<>]+(?=<\/h1>)/gm;
  const h2Re = /(?<=<h2>(<.*>)?)[^<>]+(?=<\/h2>)/gm;
  const h3Re = /(?<=<h3>(<.*>)?)[^<>]+(?=<\/h3>)/gm;
  const hRe = /(?<=<h(\d)>(<.*>)?)[^<>]+(?=<\/h(\d)>)/gm;

  const tableRegex = /(?<=<table><tr><td><p>FS ID.*<tr><td><p>)\d*(?=<\/p><\/td>)/gm;

  // const idRegexFactory = type =>
  //   new RegExp(
  //     `/(?<=<${type}>.*<table><tr><td><p>FS ID.*<tr><td><p>)\d*(?=<\/p><\/td>)/gm`
  //   );

  // const allHeadlineRegex = {
  //   h1: h1Re,
  //   h2: h2Re,
  //   h3: h3Re
  // };

  // const dataHeader = "Titles";

  class ID {
    constructor(value, position) {
      this.value = value;
      this.position = position;
    }
  }

  class Headline {
    constructor(name, type, index, pos) {
      this.name = name;
      this.type = type;
      this.index = index;
      this.pos = pos;
      this.subitems = [];
      this.idList = [];
      this.addId = id => {
        this.idList.push(id);
      };
      this.addNextIndex = index => {
        this.next = index;
      };
      this.addSubItems = subitems => {
        this.subitems = [...this.subitems, ...subitems];
      };
    }
  }

  // Create Workbook, Worksheet, set header
  const wb = new x1.Workbook();
  const ws = wb.addWorksheet("Sheet 1");
  ws.cell(1, 1)
    .string("Position")
    .style({ font: { bold: true } });
  ws.cell(1, 2)
    .string("FS ID")
    .style({ font: { bold: true } });

  ws.row(1).freeze();

  const getSourceFileName = () => {
    return new Promise((resolve, reject) => {
      fs.readdir("./source", (err, files) => {
        if (err) {
          reject(err);
        }
        resolve(files.length ? files[0].replace(".docx", "") : null);
      });
    });
  };

  const convertDocxToHtml = async fileName => {
    try {
      const result = await mammoth.convertToHtml({
        path: `./source/${fileName}.docx`
      });
      return result.value;
    } catch (e) {
      console.error(e);
    }
  };

  const getHeadlinesFromHtml = (htmlResult, type, regex) => {
    let result;
    let pos = 1;
    let headlines = [];
    while ((result = regex.exec(htmlResult)) !== null) {
      headlines[headlines.length - 1] &&
        headlines[headlines.length - 1].addNextIndex(regex.lastIndex);
      headlines.push(new Headline(result[0], type, regex.lastIndex, pos));
      pos++;
    }
    return headlines;
  };

  const getAllHeadlines = htmlResult => {
    const allHeadlines = getHeadlinesFromHtml(htmlResult, "h1", h1Re);
    allHeadlines.forEach(h1 => {
      const section = htmlResult.substring(h1.index, h1.next);
      const h2List = getHeadlinesFromHtml(section, "h2", h2Re);
      if (h2List && h2List.length) {
        h1.addSubItems(h2List);
        h2List.forEach(h2 => {
          const subSection = section.substring(h2.index, h2.next);
          const h3List = getHeadlinesFromHtml(subSection, "h3", h3Re);
          if (h3List && h3List.length) {
            h2.addSubItems(h3List);
          }
        });
      }
    });
    return allHeadlines;
  };

  const findTables = (hl, html, position) => {
    const subitems = hl.subitems;
    if (subitems.length) {
      return subitems.reduce((allIDs, subitem) => {
        const subHtmlSection = html.substring(
          subitem.index,
          subitem.next || html.length - 1
        );
        return [
          ...allIDs,
          ...findTables(subitem, subHtmlSection, `${position}.${subitem.pos}`)
        ];
      }, []);
    } else {
      const idMatches = html.match(tableRegex);
      const allIDs = idMatches
        ? idMatches.map(id => new ID(id, `${position}`))
        : [];
      return allIDs;
    }
  };

  const writeIDsToSheet = (sheet, IDs) => {
    IDs.forEach((id, index) => {
      sheet.cell(1 + index + 1, 1).string(id.position);
      sheet.cell(1 + index + 1, 2).string(id.value);
    });
  };

  const run = async () => {
    try {
      const sourceFileName = await getSourceFileName();
      if (!sourceFileName) {
        throw new Error(
          "Invalid Source File: Please make sure to put a '.docx' file into the source directory!"
        );
      }

      const htmlResult = await convertDocxToHtml(sourceFileName);
      const allHeadlines = getAllHeadlines(htmlResult);

      const allIDs = allHeadlines.reduce((allTables, hl) => {
        const tables = findTables(
          hl,
          htmlResult.substring(hl.index, hl.next || htmlResult.length - 1),
          hl.pos
        );
        return [...allTables, ...tables];
      }, []);

      // console.log("allTables", allIDs);

      writeIDsToSheet(ws, allIDs);

      // generate random String as file name to avoid file name collisions
      const randomString = `${sourceFileName}__${Math.ceil(
        Math.random() * 100000000
      )}`;

      // create results.xlsx
      if (!fs.existsSync("./results")) {
        const createResultsFolder = () =>
          new Promise((resolve, reject) => {
            fs.mkdir("./results", err => {
              reject(err);
            });
            resolve();
          });
        await createResultsFolder();
      }
      wb.write(`./results/result__${randomString}.xlsx`);
      console.log(
        `New file: 'result__${randomString}.xlsx' has successfully been created in './results/'`
      );
    } catch (err) {
      console.error(err);
    }
  };

  run();
})();
