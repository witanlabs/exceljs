const StylesXform = require('./xlsx/xform/style/styles-xform');

const ExcelJS = {
  Workbook: require('./doc/workbook'),
  ModelContainer: require('./doc/modelcontainer'),
  stream: {
    xlsx: {
      WorkbookWriter: require('./stream/xlsx/workbook-writer'),
      WorkbookReader: require('./stream/xlsx/workbook-reader'),
    },
  },
  // Style caching modes for performance optimization
  StyleCacheMode: StylesXform.StyleCacheMode,
};

Object.assign(ExcelJS, require('./doc/enums'));

module.exports = ExcelJS;
