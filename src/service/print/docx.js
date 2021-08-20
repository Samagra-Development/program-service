const { getData } = require("./dataImporter");
const docx = require("docx");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
} = docx;
// const getDocx = require("./getDocx");
const getDocx = require("./getdocxdata");
const fs = require("fs");

var {
  docDefinition,
  getStudentTeacherDetails,
  getExamName,
  getGradeHeader,
  getSubject,
  getTimeAndMarks,
  getInstructions,
  getMCQ,
  getFTB,
  getSectionTitle,
  getTF,
  getSA,
  getVSA,
  getLA,
  getComprehension,
  getMTFHeader,
  getMTFChoice,
} = require("./utils/docDefinition");
const ProgramServiceHelper = require("../../helpers/programHelper");
const axios = require("axios");
var cheerio = require("cheerio");
var cheerioTableparser = require("cheerio-tableparser");
const sizeOf = require("image-size");

const programServiceHelper = new ProgramServiceHelper();

var fonts = {
  Roboto: {
    normal: "service/print/utils/fonts/Roboto/Roboto-Regular.ttf",
    bold: "service/print/utils/fonts/Roboto/Roboto-Medium.ttf",
    italics: "service/print/utils/fonts/Roboto/Roboto-Italic.ttf",
    bolditalics: "service/print/utils/fonts/Roboto/Roboto-MediumItalic.ttf",
  },
  Hindi: {
    normal: "service/print/utils/fonts/Hindi/Jaldi-Regular.ttf",
    bold: "service/print/utils/fonts/Hindi/Jaldi-Bold.ttf",
    italics: "service/print/utils/fonts/Hindi/Jaldi-Regular.ttf",
    bolditalics: "service/print/utils/fonts/Hindi/Jaldi-Bold.ttf",
  },
  Noto: {
    normal: "service/print/utils/fonts/Noto/NotoSans-Regular.ttf",
    bold: "service/print/utils/fonts/Noto/NotoSans-SemiBold.ttf",
    italics: "service/print/utils/fonts/Noto/NotoSans-Italic.ttf",
    bolditalics: "service/print/utils/fonts/Noto/NotoSans-SemiBoldItalic.ttf",
  },
  English: {
    normal: "service/print/utils/fonts/Noto/NotoSans-Regular.ttf",
    bold: "service/print/utils/fonts/Noto/NotoSans-SemiBold.ttf",
    italics: "service/print/utils/fonts/Noto/NotoSans-Italic.ttf",
    bolditalics: "service/print/utils/fonts/Noto/NotoSans-SemiBoldItalic.ttf",
  },
};

// var printer = new PdfPrinter(fonts);
// var fs = require("fs");
const { default: Axios } = require("axios");
const { size, create } = require("lodash");

var options = {
  // ...
};

const buildDOCXWithCallback = async (id, callback) => {
  let error = false;
  let errorMsg = "";
  let totalMarks = 0;
  getData(id)
    .then(async (data) => {
      if (data.error) {
        callback(null, data.error, data.errorMsg);
      } else {
        const subject = data.paperData.subject[0];
        const grade = data.paperData.gradeLevel[0];
        const examName = data.paperData.name;
        const instructions = data.paperData.description;
        let language = data.paperData.medium[0];

        // const language = "Noto";
        data.sectionData.forEach((d) => {
          d.questions.forEach((element, index) => {
            const marks = parseInt(d.section.children[index].marks);
            if (!isNaN(marks)) totalMarks += marks;
          });
        });

        const contentBase = [
          getStudentTeacherDetails(),
          getExamName(examName),
          getGradeHeader(grade),
          getSubject(subject),
          getTimeAndMarks(90, totalMarks),
          getInstructions(instructions, language),
        ];

        const questionPaperContent = [];
        const paperDetails = {
          examName: examName,
          className: grade,
          subject: subject,
        };
        let questionCounter = 0;

        for (const d of data.sectionData) {
          // data.sectionData.forEach((d) => {
          const sectionTitle = getSectionTitle(
            d.section.name,
            detectLanguage(d.section.name)
          );
          // questionPaperContent.push(sectionTitle);
          const section = d.section;

          for (const [index, question] of d.questions.entries()) {
            questionCounter += 1;
            const marks = section.children[index].marks;
            // console.log(grade, subject, examName)
            let questionContent;
            switch (question.category) {
              case "MCQ":
                questionContent = [
                  await renderMCQ(question, questionCounter, marks),
                ];
                break;
              case "FTB":
                questionContent = [
                  await renderQuestion(question, questionCounter, marks, "FTB"),
                ];
                break;
              case "SA":
                questionContent = [
                  await renderQuestion(question, questionCounter, marks, "SA"),
                ];
                break;
              case "LA":
                questionContent = [
                  await renderQuestion(question, questionCounter, marks, "LA"),
                ];
                break;
              case "VSA":
                questionContent = [
                  await renderQuestion(question, questionCounter, marks, "VSA"),
                ];
                break;
              case "MTF":
                questionContent = await renderMTF(
                  question,
                  questionCounter,
                  marks,
                  "MTF"
                );
                break;
              case "COMPREHENSION":
                questionContent = [
                  await renderComprehension(
                    question,
                    questionCounter,
                    marks,
                    "COMPREHENSION"
                  ),
                ];
                break;
              case "CuriosityQuestion":
                questionContent = [
                  await renderQuestion(
                    question,
                    questionCounter,
                    marks,
                    "CuriosityQuestion"
                  ),
                ];
                break;
            }
            // console.log("Contents:",questionContent)
            questionPaperContent.push(questionContent);
          }
        }

        // console.log("Contents:",questionPaperContent[1])
        const doc = await getDocx.create(questionPaperContent, paperDetails);

        const b64string = await Packer.toBase64String(doc);
        let filename = grade + '_' + subject + '_' + examName
        filename = filename.replace(/\s/g, '')
        callback(b64string, null, null,filename);
      }
    })
    .catch((e) => {
      console.log(e);
      error = true;
      errorMsg = "";
      callback(null, error, errorMsg);
    });
};

const cleanHTML = (str, nbspAsLineBreak = false) => {
  // Remove HTML characters since we are not converting HTML to PDF.
  return str
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, nbspAsLineBreak ? "\n" : "");
};

const detectLanguage = (str) => {
  const unicodeBlocks = [
    {
      name: "Tamil",
      regex: /[\u0B80-\u0BFF]+/g,
    },
    {
      name: "Hindi",
      regex: /[\u0900-\u097F]+/g,
    },
  ];

  let language = "English";

  const langSplit = {
    Hindi: 0,
    Tamil: 0,
    English: 0,
    Undefined: 0,
  };
  if (typeof str === "string") {
    str.split("").forEach((letter) => {
      let found = false;
      unicodeBlocks.forEach((block) => {
        if (letter.match(block.regex)) {
          langSplit[block.name]++;
          found = true;
        }
      });
      if (!found) {
        langSplit.English++;
      }
    });

    let max = 0;
    for (var key of Object.keys(langSplit)) {
      if (langSplit[key] > max) {
        max = langSplit[key];
        language = key;
      }
    }

    return language;
  }
  return "English";
};

function createImageElement(src, width) {
  let imageElement = {};
  if (src.search("image/gif") >= 0) return null;
  imageElement.image = src;
  let img = Buffer.from(src.split(";base64,").pop(), "base64");
  let dimensions = sizeOf(img);
  console.log("Dimensions", dimensions);
  let resizedWidth = dimensions.width * width;
  let resizedHeight = dimensions.height * width;
  imageElement.width = resizedWidth > 150 ? 150 : resizedWidth;
  imageElement.height = resizedHeight > 150 ? 150 : resizedHeight;
  return imageElement;
}

function extractTextFromElement(elem) {
  let rollUp = "";
  if (cheerio.text(elem)) return cheerio.text(elem);
  else if (elem.name === "sup")
    return { text: elem.children[0].data, superScript: true };
  else if (elem.name === "sub")
    return { text: elem.children[0].data, subScript: true };
  else if (elem.type === "text" && elem.data) return elem.data;
  else {
    if (elem.children && elem.children.length) {
      for (const nestedElem of elem.children) {
        let recurse = extractTextFromElement(nestedElem);
        if (Array.isArray(rollUp)) {
          rollUp.push(recurse);
        } else {
          if (Array.isArray(recurse)) {
            rollUp = recurse;
          } else if (typeof recurse === "object") {
            rollUp = [rollUp, recurse];
          } else rollUp += recurse;
        }
      }
    }
  }
  return rollUp;
}

async function getStack(htmlString, questionCounter) {
  const stack = [];
  $ = cheerio.load(htmlString);
  const elems = $("body").children().toArray();
  for (const [index, elem] of elems.entries()) {
    let nextLine = "";
    switch (elem.name) {
      case "p":
        let extractedText = extractTextFromElement(elem);
        // Returns array if superscript/subscript inside
        if (Array.isArray(extractedText)) nextLine = { text: extractedText };
        else nextLine += extractedText;
        break;
      case "ol":
        nextLine = {
          ol: elem.children.map(
            (el) =>
              el.children[0] &&
              (el.children[0].data ||
                (el.children[0].children[0] && el.children[0].children[0].data))
          ),
        };
        break;
      case "ul":
        nextLine = {
          ul: elem.children.map(
            (el) =>
              el.children[0] &&
              (el.children[0].data ||
                (el.children[0].children[0] && el.children[0].children[0].data))
          ),
        };
        break;
      case "figure":
        let { style } = elem.attribs;
        // console.log("Style:",style)
        let width = 1;
        if (style) {
          width = parseFloat(style.split(":").pop().slice(0, -2));
          width = width / 100;
        }
        if (elem.children && elem.children.length) {
          let { src } = elem.children[0].attribs;
          if (src) {
            switch (src.slice(0, 4)) {
              case "data":
                nextLine = createImageElement(src, width);
                break;
              default:
                let res = await programServiceHelper.getQuestionMedia(src);
                nextLine = createImageElement(res, width);
            }
          }
        }
        if (!nextLine)
          nextLine = "<An image of an unsupported format was scrubbed>";
        break;
    }
    if (index === 0 && questionCounter) {
      if (elem.name === "p") {
        if (typeof nextLine === "object")
          nextLine = { text: [`${questionCounter}. `, ...nextLine.text] };
        else nextLine = `${questionCounter}. ${nextLine}`;
      } else stack.push(`${questionCounter}.`);
    }
    stack.push(nextLine);
  }
  return stack;
}

async function renderMCQ(question, questionCounter, marks) {
  const questionOptions = [],
    answerOptions = ["A", "B", "C", "D"];
  let questionTitle;
  for (const [index, qo] of question.editorState.options.entries()) {
    let qoBody = qo.value.body;
    let qoData =
      qoBody.search("img") >= 0 ||
      qoBody.search("sup") >= 0 ||
      qoBody.search("sub") >= 0 ||
      qoBody.match(/<p>/g).length > 1
        ? await getStack(qoBody, answerOptions[index])
        : [`${answerOptions[index]}. ${cleanHTML(qoBody)}`];
    questionOptions.push(qoData);
  }
  let q = question.editorState.question;
  questionTitle =
    q.search("img") >= 0 ||
    q.search("sub") >= 0 ||
    q.search("sup") >= 0 ||
    q.match(/<p>/g).length > 1
      ? await getStack(q, questionCounter)
      : [`${questionCounter}. ${cleanHTML(q)}`];

  let questionOpt = [];
  let imageProperties = [];
  if (typeof questionOptions[0][1] === "object") {
    // console.log("Que:",questionOptions)
    questionOpt.push(questionOptions[0][0] + questionOptions[0][1].image);
    imageProperties.push({
      width: questionOptions[0][1].width,
      height: questionOptions[0][1].height,
    });
  } else {
    questionOpt.push(questionOptions[0][0]);
    imageProperties.push({
      width: 0,
      height: 0,
    });
  }

  if (typeof questionOptions[1][1] === "object") {
    questionOpt.push(questionOptions[1][0] + questionOptions[1][1].image);
    imageProperties.push({
      width: questionOptions[1][1].width,
      height: questionOptions[1][1].height,
    });
  } else {
    questionOpt.push(questionOptions[1][0]);
    imageProperties.push({
      width: 0,
      height: 0,
    });
  }

  if (typeof questionOptions[2][1] === "object") {
    questionOpt.push(questionOptions[2][0] + questionOptions[2][1].image);
    imageProperties.push({
      width: questionOptions[2][1].width,
      height: questionOptions[2][1].height,
    });
  } else {
    questionOpt.push(questionOptions[2][0]);
    imageProperties.push({
      width: 0,
      height: 0,
    });
  }

  if (typeof questionOptions[3][1] === "object") {
    questionOpt.push(questionOptions[3][0] + questionOptions[3][1].image);
    imageProperties.push({
      width: questionOptions[3][1].width,
      height: questionOptions[3][1].height,
    });
  } else {
    questionOpt.push(questionOptions[3][0]);
    imageProperties.push({
      width: 0,
      height: 0,
    });
  }

  let data = {
    Questions: questionTitle,
    Option1: questionOpt[0],
    Option2: questionOpt[1],
    Option3: questionOpt[2],
    Option4: questionOpt[3],
    Marks: marks,
    Language: detectLanguage(questionTitle[0]),
    type: "MCQ",
    height: imageProperties[0].height,
    width: imageProperties[0].width,
  };
  return data;
  
}

async function renderQuestion(question, questionCounter, marks, Type) {
  let data;
  if (
    (question.media && question.media.length) ||
    question.editorState.question.search("img") >= 0 ||
    question.editorState.question.search("sub") >= 0 ||
    question.editorState.question.search("sup") >= 0 ||
    question.editorState.question.search("ol") >= 0 ||
    question.editorState.question.search("ul") >= 0 ||
    question.editorState.question.match(/<p>/g).length > 1
  ) {
    data = await getStack(question.editorState.question, questionCounter);
  } else {
    data = [`${questionCounter}. ${cleanHTML(question.editorState.question)}`];
  }
  
  let quedata = {
    Questions: data,
    Marks: marks,
    type: Type,
  };
 
  return quedata;
}

async function renderComprehension(question, questionCounter, marks, Type) {
  let data;
  if (
    (question.media && question.media.length) ||
    question.editorState.question.search("img") >= 0 ||
    question.editorState.question.search("sub") >= 0 ||
    question.editorState.question.search("sup") >= 0 ||
    question.editorState.question.search("ol") >= 0 ||
    question.editorState.question.search("ul") >= 0 ||
    question.editorState.question.match(/<p>/g).length > 1
  ) {
    data = await getStack(question.editorState.question, questionCounter);
  } else {
    data = [`${questionCounter}. ${cleanHTML(question.editorState.question)}`];
  }
  let quedata = {
    Questions: data,
    Marks: marks,
    type: Type,
  };
  return quedata;
  // return getComprehension(data, detectLanguage(data[0]), marks);
}

function renderTF(question, questionCounter, marks) {
  const questionTitle =
    questionCounter + ". " + cleanHTML(question.editorState.question);
  return getTF(questionTitle, detectLanguage(questionTitle[0]), marks);
}

async function renderMTF(question, questionCounter, marks, Type) {
  $ = cheerio.load(question.editorState.question);
  cheerioTableparser($);
  var data = [];
  var columns = $("table").parsetable(false, false, false);
  let transposeColumns = columns[0].map((_, colIndex) =>
    columns.map((row) => row[colIndex])
  );

  const heading = questionCounter + ". " + cleanHTML($("p").first().text());
  console.log("MTF:", heading);
  data.push(heading, detectLanguage(heading), marks);
  // console.log("MTF:",transposeColumns)
  data.push(
    getMTFHeader(
      cleanHTML(transposeColumns[0][0]),
      cleanHTML(transposeColumns[0][1]),
      detectLanguage(cleanHTML(transposeColumns[0][0]))
    )
  );

  transposeColumns.shift();

  const rows = [];
  for (const r of transposeColumns) {
    let left, right;
    if (r[0].search("img") >= 0) {
      left = await getStack(r[0]);
    } else left = [cleanHTML(r[0])];
    if (r[1].search("img") >= 0) {
      right = await getStack(r[1]);
    } else right = [cleanHTML(r[1])];
    rows.push({left, right});
    
  }

  // return data.concat(rows);
  let mtfData = [{
    Questions: rows,
    Marks: marks,
    type: Type,
    heading:heading
  }];
  return mtfData
}
module.exports = {
  buildDOCXWithCallback,
};
