const _ = require('../../../utils/under-dash');

const colCache = require('../../../utils/col-cache');
const XmlStream = require('../../../utils/xml-stream');
const xmlUtils = require('../../../utils/utils');

const BaseXform = require('../base-xform');
const StaticXform = require('../static-xform');
const ListXform = require('../list-xform');
const DefinedNameXform = require('./defined-name-xform');
const SheetXform = require('./sheet-xform');
const WorkbookViewXform = require('./workbook-view-xform');
const WorkbookPropertiesXform = require('./workbook-properties-xform');
const WorkbookCalcPropertiesXform = require('./workbook-calc-properties-xform');
const WorkbookPivotCacheXform = require('./workbook-pivot-cache-xform');

function xmlEncodeAttr(str) {
  return xmlUtils.xmlEncode(String(str));
}

class WorkbookXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      fileVersion: WorkbookXform.STATIC_XFORMS.fileVersion,
      workbookPr: new WorkbookPropertiesXform(),
      bookViews: new ListXform({
        tag: 'bookViews',
        count: false,
        childXform: new WorkbookViewXform(),
      }),
      sheets: new ListXform({tag: 'sheets', count: false, childXform: new SheetXform()}),
      definedNames: new ListXform({
        tag: 'definedNames',
        count: false,
        childXform: new DefinedNameXform(),
      }),
      calcPr: new WorkbookCalcPropertiesXform(),
      pivotCaches: new ListXform({
        tag: 'pivotCaches',
        count: false,
        childXform: new WorkbookPivotCacheXform(),
      }),
    };
  }

  prepare(model) {
    model.sheets = model.worksheets;

    // collate all the print areas from all of the sheets and add them to the defined names
    const printAreas = [];
    let index = 0; // sheets is sparse array - calc index manually
    model.sheets.forEach(sheet => {
      if (sheet.pageSetup && sheet.pageSetup.printArea) {
        const definedName = {
          name: '_xlnm.Print_Area',
          ranges: sheet.pageSetup.printArea.split('&&').map(printArea => {
            const printAreaComponents = printArea.split(':');
            // 既に $ 付きの場合はそのまま使う、なければ付与する
            const tl = printAreaComponents[0].startsWith('$') ? printAreaComponents[0] : `$${printAreaComponents[0]}`;
            const br = printAreaComponents[1] && printAreaComponents[1].startsWith('$') ? printAreaComponents[1] : `$${printAreaComponents[1] || ''}`;
            // シート名にスペースや特殊文字が含まれる場合のみシングルクォートで囲む
            const needsQuote = /[\s\(\)\[\]'!]/.test(sheet.name);
            const sheetRef = needsQuote ? `'${sheet.name}'` : sheet.name;
            return [`${sheetRef}!${tl}:${br}`];
          }),
          localSheetId: index,
        };
        printAreas.push(definedName);
      }

      if (sheet.pageSetup && (sheet.pageSetup.printTitlesRow || sheet.pageSetup.printTitlesColumn)) {
        const ranges = [];
        const needsQuoteTitles = /[\s\(\)\[\]'!]/.test(sheet.name);
        const sheetRefTitles = needsQuoteTitles ? `'${sheet.name}'` : sheet.name;

        if (sheet.pageSetup.printTitlesColumn) {
          const titlesColumns = sheet.pageSetup.printTitlesColumn.split(':');
          ranges.push(`${sheetRefTitles}!$${titlesColumns[0]}:$${titlesColumns[1]}`);
        }

        if (sheet.pageSetup.printTitlesRow) {
          const titlesRows = sheet.pageSetup.printTitlesRow.split(':');
          ranges.push(`${sheetRefTitles}!$${titlesRows[0]}:$${titlesRows[1]}`);
        }

        const definedName = {
          name: '_xlnm.Print_Titles',
          ranges,
          localSheetId: index,
        };

        printAreas.push(definedName);
      }
      index++;
    });
    if (printAreas.length) {
      model.definedNames = model.definedNames.concat(printAreas);
    }

    (model.media || []).forEach((medium, i) => {
      // assign name
      medium.name = medium.type + (i + 1);
    });
  }

  render(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('workbook', WorkbookXform.WORKBOOK_ATTRIBUTES);

    this.map.fileVersion.render(xmlStream);
    this.map.workbookPr.render(xmlStream, model.properties);
    this.map.bookViews.render(xmlStream, model.views);
    this.map.sheets.render(xmlStream, model.sheets);
    // externalReferences を raw XML として出力
    if (model.externalReferencesXml) {
      xmlStream.writeXml(model.externalReferencesXml);
    }
    this.map.definedNames.render(xmlStream, model.definedNames);
    this.map.calcPr.render(xmlStream, model.calcProperties);
    this.map.pivotCaches.render(xmlStream, model.pivotTables);

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'workbook':
        this._extRefDepth = 0;
        this._extRefParts = null;
        return true;
      case 'externalReferences':
        this._extRefDepth = 1;
        this._extRefParts = ['<externalReferences>'];
        return true;
      default:
        if (this._extRefDepth > 0) {
          this._extRefDepth++;
          const attrsStr = Object.entries(node.attributes || {})
            .map(([k, v]) => ` ${k}="${xmlEncodeAttr(v)}"`)
            .join('');
          this._extRefParts.push(`<${node.name}${attrsStr}>`);
          return true;
        }
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
    }
  }

  parseText(text) {
    if (this._extRefDepth > 0) {
      if (text) this._extRefParts.push(text);
      return;
    }
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this._extRefDepth > 0) {
      if (name === 'externalReferences') {
        this._extRefDepth = 0;
        this._extRefParts.push('</externalReferences>');
        this._externalReferencesXml = this._extRefParts.join('');
        this._extRefParts = null;
      } else {
        this._extRefDepth--;
        // 自己閉じタグかどうかに関わらず閉じタグを追加
        this._extRefParts.push(`</${name}>`);
      }
      return true;
    }
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'workbook':
        this.model = {
          sheets: this.map.sheets.model,
          properties: this.map.workbookPr.model || {},
          views: this.map.bookViews.model,
          calcProperties: {},
        };
        if (this.map.definedNames.model) {
          this.model.definedNames = this.map.definedNames.model;
        }
        // externalReferences を raw XML として保存
        if (this._externalReferencesXml) {
          this.model.externalReferencesXml = this._externalReferencesXml;
        }

        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }

  reconcile(model) {
    const rels = (model.workbookRels || []).reduce((map, rel) => {
      map[rel.Id] = rel;
      return map;
    }, {});

    // reconcile sheet ids, rIds and names
    const worksheets = [];
    let worksheet;
    let index = 0;

    (model.sheets || []).forEach(sheet => {
      const rel = rels[sheet.rId];
      if (!rel) {
        return;
      }
      // if rel.Target start with `[space]/xl/` or `/xl/` , then it will be replaced with `''` and spliced behind `xl/`,
      // otherwise it will be spliced directly behind `xl/`. i.g.
      worksheet = model.worksheetHash[`xl/${rel.Target.replace(/^(\s|\/xl\/)+/, '')}`];
      // If there are "chartsheets" in the file, rel.Target will
      // come out as chartsheets/sheet1.xml or similar here, and
      // that won't be in model.worksheetHash.
      // As we don't have the infrastructure to support chartsheets,
      // we will ignore them for now:
      if (worksheet) {
        worksheet.name = sheet.name;
        worksheet.id = sheet.id;
        worksheet.state = sheet.state;
        worksheets[index++] = worksheet;
      }
    });

    // reconcile print areas
    const definedNames = [];
    _.each(model.definedNames, definedName => {
      if (definedName.name === '_xlnm.Print_Area') {
        worksheet = worksheets[definedName.localSheetId];
        if (worksheet) {
          if (!worksheet.pageSetup) {
            worksheet.pageSetup = {};
          }
          let printAreaStr;
          try {
            const range = colCache.decodeEx(definedName.ranges[0]);
            // $A$1:$BL$624 形式で保存する
            if (range.tl && range.br) {
              printAreaStr = `${range.tl['$col$row']}:${range.br['$col$row']}`;
            } else {
              printAreaStr = range.dimensions;
            }
          } catch (e) {
            printAreaStr = definedName.ranges[0];
          }
          worksheet.pageSetup.printArea = worksheet.pageSetup.printArea
            ? `${worksheet.pageSetup.printArea}&&${printAreaStr}`
            : printAreaStr;
        }
      } else if (definedName.name === '_xlnm.Print_Titles') {
        worksheet = worksheets[definedName.localSheetId];
        if (worksheet) {
          if (!worksheet.pageSetup) {
            worksheet.pageSetup = {};
          }

          const rangeString = definedName.ranges.join(',');

          const dollarRegex = /\$/g;

          const rowRangeRegex = /\$\d+:\$\d+/;
          const rowRangeMatches = rangeString.match(rowRangeRegex);

          if (rowRangeMatches && rowRangeMatches.length) {
            const range = rowRangeMatches[0];
            worksheet.pageSetup.printTitlesRow = range.replace(dollarRegex, '');
          }

          const columnRangeRegex = /\$[A-Z]+:\$[A-Z]+/;
          const columnRangeMatches = rangeString.match(columnRangeRegex);

          if (columnRangeMatches && columnRangeMatches.length) {
            const range = columnRangeMatches[0];
            worksheet.pageSetup.printTitlesColumn = range.replace(dollarRegex, '');
          }
        }
      } else {
        definedNames.push(definedName);
      }
    });
    model.definedNames = definedNames;

    // used by sheets to build their image models
    model.media.forEach((media, i) => {
      media.index = i;
    });
  }
}

WorkbookXform.WORKBOOK_ATTRIBUTES = {
  xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  'mc:Ignorable': 'x15',
  'xmlns:x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
};
WorkbookXform.STATIC_XFORMS = {
  fileVersion: new StaticXform({
    tag: 'fileVersion',
    $: {appName: 'xl', lastEdited: 5, lowestEdited: 5, rupBuild: 9303},
  }),
};

module.exports = WorkbookXform;
