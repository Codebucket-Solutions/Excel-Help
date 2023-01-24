let xlsx = require("xlsx");

let XlsxBuilder = function () {
  let workbook = xlsx.utils.book_new();

  return {
    XLSX: xlsx,
    addJson: function (ws, json, config = {}) {
      if (!config) config = {};
      let { headerMap = {}, excludeHeaders = [] } = config;
      let keyOperations = function (json, newKeys = {}, excludeHeaders = []) {
        json.map((obj) => {
          const keyValues = Object.keys(obj).map((key) => {
            const newKey = newKeys[key] || key;
            if (!excludeHeaders.includes(newKey)) return { [newKey]: obj[key] };
          });
          return Object.assign({}, ...keyValues);
        });
        return json;
      };
      json = keyOperations(json, headerMap, excludeHeaders);
      if (!config.origin) config.origin = -1;
      return xlsx.utils.sheet_add_json(ws, json, config);
    },

    addAoa: function (ws, aoa, config = {}) {
      if (!config) config = {};
      if (!config.origin) config.origin = -1;
      return xlsx.utils.sheet_add_aoa(ws, aoa, config);
    },

    addSheet: function (values, sheetName, config = {}) {
      let ws = null;
      for (let value of values) {
        if (value.type == "json") {
          ws = this.addJson(ws, value.data, value.options);
        } else if (value.type == "columns") {
          ws = this.addAoa(ws, value.data, value.options);
        }
      }
      if (config) {
        if (Object.keys(config).length) {
          for (let key of Object.keys(config)) {
            ws[key] = config[key];
          }
        }
      }
      xlsx.utils.book_append_sheet(workbook, ws, sheetName);
      return this;
    },
    build: function () {
      return workbook;
    },
    export: function (options) {
      return xlsx.write(workbook, options);
    },
  };
};

XlsxBuilder.middleware = function (req, res, next) {
  res.xlsx = function (fn, data, config) {
    if (!config) config = { type: "buffer", bookType: "xlsx" };
    if (!config.type) config.type = "buffer";
    if (!config.bookType) config.bookType = "xlsx";
    let xlsxWb = new XlsxBuilder().addSheet(data).export(config);
    res.setHeader("Content-Type", "application/vnd.ms-excel");
    res.setHeader("Content-Disposition", "attachment; filename=" + fn);
    res.end(xlsxWb, "binary");
  };
  next();
};

module.exports = XlsxBuilder;
