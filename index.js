(async () => {
  const mammoth = require("mammoth");
  const x1 = require("excel4node");
  const fs = require("fs");

  const tableRegex = /(?<=<table><tr><td><p>FS ID.*<tr><td><p>)\d*\s*(?=<\/p><\/td>)/gm;

  const headlineRegexFactory = type =>
    new RegExp(
      `(?<=<${type}>(<\/?[^hH].*>)?)[^<>]+(?=(<[^hH]*>)?<\/${type}>)`,
      "gm"
    );

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
  ws.cell(1, 4)
    .string("Result")
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

  const getSubHeadlines = (hl, html, type) => {
    const hlType = `h${type}`;
    const regex = headlineRegexFactory(hlType);
    // console.log("regex", regex);
    const subHeadlines = getHeadlinesFromHtml(html, hlType, regex);
    if (subHeadlines && subHeadlines.length) {
      hl.addSubItems(subHeadlines);
      hl.subitems.forEach(sHl => {
        const section = html.substring(sHl.index, sHl.next || html.length - 1);
        getSubHeadlines(sHl, section, type + 1);
      });
    }
  };

  const getAllHeadlines = htmlResult => {
    const headlineType = 1;
    const hlType = `h${headlineType}`;
    const regex = headlineRegexFactory(hlType);
    // console.log("regex", regex);
    const headlines = getHeadlinesFromHtml(htmlResult, hlType, regex);
    if (headlines && headlines.length) {
      headlines.forEach(hl => {
        const section = htmlResult.substring(hl.index, hl.next);
        getSubHeadlines(hl, section, headlineType + 1);
      });
    }
    return headlines;
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
      // console.log("idMatches", idMatches, position);
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
      sheet.cell(1 + index + 1, 4).string(`${id.position} - ${id.value}`);
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

      // console.log("allHeadlines[4].subitems[1]", allHeadlines[4].subitems[1]);
      const allIDs = allHeadlines.reduce((allTables, hl) => {
        const tables = findTables(
          hl,
          htmlResult.substring(hl.index, hl.next || htmlResult.length - 1),
          hl.pos
        );
        return [...allTables, ...tables];
      }, []);

      // console.log("allIDs", allIDs);

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
