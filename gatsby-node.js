const Sheets = require("node-sheets").default;
const createNodeHelpers = require("gatsby-node-helpers").default;
const camelCase = require("camelcase");

exports.sourceNodes = async ({ actions, createNodeId }, pluginOptions) => {
  const { createNode } = actions;
  const {
    spreadsheetId,
    spreadsheetName = "",
    typePrefix = "GoogleSpreadsheet",
    credentials,
    filterNode = () => true,
    mapNode = (node) => node,
  } = pluginOptions;

  const { createNodeFactory } = createNodeHelpers({
    typePrefix,
  });

  const gs = new Sheets(spreadsheetId);

  if (credentials) {
    await gs.authorizeJWT(credentials);
  }

  const promises = (await gs.getSheetsNames()).map(async (sheetTitle) => {
    const tables = await gs.tables(sheetTitle);
    const { rows, formats, headers } = tables;

    const buildNode = createNodeFactory(
      camelCase(`${spreadsheetName} ${sheetTitle}`)
    );

    rows
      .map((row) => toNode(row, sheetTitle))
      .filter(filterNode)
      .map(mapNode)
      .forEach((node, i) => {
        const hasProperties = Object.values(node).some(
          (value) => value !== null
        );
        if (hasProperties) {
          createNode({
            ...buildNode(node),
            id: createNodeId(
              `${typePrefix} ${spreadsheetName} ${sheetTitle} ${i}`
            ),
          });
        }
      });
  });
  return Promise.all(promises);
};

function toNode(row, sheetTitle) {
  return Object.entries(row).reduce((obj, [key, cell]) => {
    if (key === undefined || key === "undefined") {
      return obj;
    }

    // `node-sheets` adds default values for missing numbers and dates, by checking
    // for the precense of `stringValue` (the formatted value), we can ensure that
    // the value actually exists.
    const value =
      typeof cell === "object" && cell.stringValue !== undefined
        ? cell.value
        : null;
    obj[camelCase(key)] = value;

    if (cell && cell.textFormatRuns) {
      let htmlContent = "";
      htmlContent = convertToHTML(cell, value);
      console.log("sheet: ", sheetTitle);
      console.log("textFormatRuns");
      console.log(JSON.stringify(cell.textFormatRuns, null, 2));
      console.log(`htmlContent`, htmlContent);
      obj.htmlContent = htmlContent;
    }

    return obj;
  }, {});
}

/**
  * Add single tag only
  */
const addTag = ({
  tag,
  attribute,
  isStarted,
  value,
  htmlContent,
  prev,
  curr,
}) => {

  // Add sinle tag only
  if (Object.values(curr.format).length > 1) return htmlContent;


  if (curr.format[attribute]) {
    // add start tag
    if (prev && !isStarted) {
      const normalText = value.substring(prev.startIndex, curr.startIndex);
      htmlContent = `${htmlContent}${normalText}<${tag}>`;
    } else {
      htmlContent = `<${tag}>`;
    }
  } else {
    // add end tag
    if (prev && isStarted) {
      const text = value.substring(prev.startIndex || 0, curr.startIndex);
      htmlContent = `${htmlContent}${text}</${tag}>`;
    }
  }
  return htmlContent;
};

const convertToHTML = (cell, value) => {
  let htmlContent = "";
  let prev;

  let isAlreadyBold = false;
  let isAlreadyItalic = false;
  let isAlreadyUnderline = false;
  let isAlreadyStrikethrough = false;

  cell.textFormatRuns.forEach((curr, idx, arr) => {
    htmlContent = addTag({
      tag: "strong",
      attribute: 'bold',
      isStarted: isAlreadyBold,
      value,
      htmlContent,
      prev,
      curr,
    });

    htmlContent = addTag({
      tag: "em",
      attribute: 'italic',
      isStarted: isAlreadyItalic,
      value,
      htmlContent,
      prev,
      curr,
    });

    htmlContent = addTag({
      tag: "u",
      attribute: 'underline',
      isStarted: isAlreadyUnderline,
      value,
      htmlContent,
      prev,
      curr,
    });

    htmlContent = addTag({
      tag: "del",
      attribute: 'strikethrough',
      isStarted: isAlreadyStrikethrough,
      value,
      htmlContent,
      prev,
      curr,
    });


    // last item
    if (idx === arr.length - 1) {
      const rest = value.substring(curr.startIndex);
      htmlContent = `${htmlContent}${rest}`;
    }

    prev = curr;
    isAlreadyBold = !!curr.format.bold;
    isAlreadyItalic = !!curr.format.italic;
    isAlreadyUnderline = !!curr.format.underline;
    isAlreadyStrikethrough = !!curr.format.strikethrough;
  });

  return htmlContent;
};
