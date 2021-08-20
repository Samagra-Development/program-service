const docx = require("docx");
const { text } = require("express");
const fs = require("fs");
const {
  Document,
  Packer,
  BorderStyle,
  Paragraph,
  TextRun,
  AlignmentType,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  HeadingLevel,
} = docx;

function create(data, paperData) {
  // console.log("Data:", data)
  let queNum = 0;
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          headers(),
          new Paragraph({
            children: [], // Just newline without text
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `${paperData.examName}`,
                heading: HeadingLevel.TITLE,
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [], // Just newline without text
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `Grade: ${paperData.className}`,
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [], // Just newline without text
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `Subject: ${paperData.subject}`,
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [], // Just newline without text
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "__________________________________________________________________________________________",
                bold: true,
                thematicBreak: true,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({
                text: "List of questions",
                italics: true,
                bold: true,
                thematicBreak: true,
              }),
            ],
          }),

          ...data
            .map((question) => {
              const arr = [];

              if (question !== undefined) {
                let page = 1;
                if (question[0].type === "COMPREHENSION") {
                  let count = 0;
                  arr.push(Marks(question));
                  question[0].Questions.map((item) => {
                      if (item.ol) {
                        let count1 = 0;
                        item.ol.map((text) => {
                          count1++;
                          if (typeof text === "object") {
                            text = count1 + text.text;
                            arr.push(
                              new Paragraph({
                                text: `${count1}.${text}`,
                              })
                            );
                          } else {
                            arr.push(
                              new Paragraph({
                                children: [
                                  new TextRun({
                                    text: `${count1}.${text}`,
                                  }),
                                ],
                              })
                            );
                          }
                        });
                    }
                    arr.push(createCOMPREHENSIONObject(item, count++));
                  });

                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "CuriosityQuestion") {
                  // console.log("CuriosityQuestion Dta:", question[0].Questions);
                } else if (question[0].type === "SA") {
                  let count = 0;
                  arr.push(Marks(question));
                  question[0].Questions.map((item) => {
                    arr.push(createSAObject(item, count++));
                  });

                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "LA") {
                  // console.log("LA Dta:", page);
                  arr.push(Marks(question));
                  question[0].Questions.map((item) => {
                    arr.push(createSAObject(item));
                  });
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "VSA") {
                  // console.log("VSA Dta:", question[0].Questions[0]);
                  let count = 0;
                  arr.push(Marks(question));
                  question[0].Questions.map((item) => {
                    // console.log(item)

                    arr.push(createSAObject(item, count++));
                  });
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "FTB") {
                  let count = 0;
                  arr.push(Marks(question));
                  question[0].Questions.map((item) => {
                    if (typeof item === "object") {
                      arr.push(createFTBObject(item));
                    } else {
                      arr.push(createFTB(item, count++));
                    }
                  });
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "MCQ") {
                  let testimage = formatOptions(question[0]);
                  arr.push(Marks(question));
                  let count = 0;
                  question[0].Questions.map((item) => {
                    arr.push(createSAObject(item));
                  });

                  arr.push(optionsTabel(testimage));
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "MTF") {
                  arr.push(Marks(question));

                  let count = 0;
                  arr.push(createSAObject(question[0].heading, count));
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                  arr.push(MTFTabel(question));
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                }
              }
              return arr;
            })
            .reduce((prev, curr) => prev.concat(curr), []),
        ],
      },
    ],
  });
  return doc;
}
function MTFTabel(question) {
  const arr = [];
  const rowheading = new TableRow({
    children: [
      new TableCell({
        borders: MTFborder,
        width: {
          size: 5505,
          type: WidthType.DXA,
        },
        children: [
          new Paragraph({
            text: "Column1",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        borders: MTFborder,
        width: {
          size: 5505,
          type: WidthType.DXA,
        },
        children: [
          new Paragraph({
            text: "Column2",
            bold: true,
          }),
        ],
      }),
    ],
  });
  arr.push(rowheading);
  question[0].Questions.map((item) => {
    arr.push(
      new TableRow({
        children: [
          new TableCell({
            borders: MTFborder,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [createSAObject(item.left[0], 0)],
          }),
          new TableCell({
            borders: MTFborder,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [createSAObject(item.right[0], 0)],
          }),
        ],
      })
    );
  });
  // console.log("Arr:",arr)
  //
  return new Table({
    columnWidths: [5505, 5505],
    rows: arr,
  });
}
function Marks(data) {
  if (data[0].Marks !== undefined) {
    return new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [
        new TextRun({
          text: `${data[0].Marks}`,
          bold: true,
        }),
      ],
    });
  } else {
    return new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [],
    });
  }
}

function createFTBObject(data) {
  const arr = [];

  data.text
    .map((text) => {
      if (typeof text === "object") {
        arr.push(new TextRun(text));
      } else {
        arr.push(
          new TextRun({
            text: `${text}`,
          })
        );
      }
    })
    .reduce((prev, curr) => prev.concat(curr), []);
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    children: arr,
  });
}

function createFTB(data, count) {
  // console.log("Data:",data, count)
  if (count !== 0) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({
          text: "    " + `${data}`,
          thematicBreak: true,
        }),
      ],
    });
  } else {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({
          text: `${data}`,
          thematicBreak: true,
        }),
      ],
    });
  }
}

function createSAObject(data, count) {
  // console.log("Item:", data)
  const arr = [];
  if (data.text) {
    data.text
      .map((text) => {
        if (typeof text === "object") {
          arr.push(new TextRun(text));
        } else {
          arr.push(
            new TextRun({
              text: `${text}`,
            })
          );
        }
      })
      .reduce((prev, curr) => prev.concat(curr), []);
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: arr,
    });
  } else if (data.image) {
    if (data.image.includes("data:image/")) {
      let image = getBufferImg(data.image);

      return new Paragraph({
        children: [
          new ImageRun({
            data: image,
            transformation: {
              width: data.width,
              height: data.height,
            },
          }),
        ],
      });
    }
  } else {
    return createFTB(data, count);
  }
}

function createCOMPREHENSIONObject(data, count) {
  // console.log("Data:",data)
  const arr = [];
  if (data.text) {
    data.text
      .map((text) => {
        if (typeof text === "object") {
          arr.push(new TextRun(text));
        } else {
          arr.push(
            new TextRun({
              text: `${text}`,
            })
          );
        }
      })
      .reduce((prev, curr) => prev.concat(curr), []);
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: arr,
    });
  } else if (data.image) {
    if (data.image.includes("data:image/")) {
      let image = getBufferImg(data.image);

      return new Paragraph({
        children: [
          new ImageRun({
            data: image,
            transformation: {
              width: data.width,
              height: data.height,
            },
          }),
        ],
      });
    }
  } else if (data.ol) {
    // let count1 = 0;
    // data.ol
    //   .map((text) => {
    //     count1++;
    //     if (typeof text === "object") {
    //       text = count1 + text.text;
    //       arr.push(
    //         new Paragraph({
    //           text: `${count1}.${text}`,
    //         })
    //       );
    //     } else {
    //       arr.push(
    //         new Paragraph({
    //           // children: [
    //           //   new TextRun({
    //               text: `${count1}.${text}`,
    //           //   }),
    //           // ],
    //         })
    //       );
    //     }
    //   })
    //   .reduce((prev, curr) => prev.concat(curr), []);
    //   console.log("ol",arr)
    // return new Paragraph({
    //   alignment: AlignmentType.LEFT,
    //   children: arr,
    // });
  } else {
    return createFTB(data, count);
  }
}

function imageData(image) {
  let bufferImage;
  if (image.includes("data:image/png;base64,")) {
    return (bufferImage = image.replace("data:image/png;base64,", ""));
  } else if (image.includes("data:image/jpg;base64")) {
    return (bufferImage = image.replace("data:image/jpg;base64,", ""));
  } else if (image.includes("data:image/jpeg;base64")) {
    return (bufferImage = image.replace("data:image/jpeg;base64,", ""));
  }
}
function getBufferData(data) {
  let image = imageData(data);
  return image.substr(2);
}
function getBufferImg(data) {
  let image = imageData(data);
  return image;
}
function formatOptions(data) {
  // console.log("Options:",data)
  let optionArr = [];
  let image;
  let testimage = data;
  if (testimage) {
    if (testimage.Option1.includes("data:image/")) {
      optionArr.push(testimage.Option1);
    } else {
      optionArr.push(testimage.Option1);
    }
    if (testimage.Option2.includes("data:image/")) {
      optionArr.push(testimage.Option2);
    } else {
      optionArr.push(testimage.Option2);
    }
    if (testimage.Option3.includes("data:image/")) {
      optionArr.push(testimage.Option3);
    } else {
      optionArr.push(testimage.Option3);
    }
    if (testimage.Option4.includes("data:image/")) {
      optionArr.push(testimage.Option4);
    } else {
      optionArr.push(testimage.Option4);
    }
    optionArr.push(testimage.height);
    optionArr.push(testimage.width);
  }
  // console.log("options :",optionArr)
  return optionArr;
}

function displayOptions(option, height, width) {
  // console.log("image", option);
  if (option.includes("data:image/")) {
    let image = getBufferData(option);
      return new Paragraph({
      text: option.substr(0, 1),
      children: [
        new ImageRun({
          data: image,
          transformation: {
            width: width,
            height: height,
          },
        }),
      ],
    });
  } else {
    // console.log("Text")
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({
          text: option,
        }),
      ],
    });
  }
}

const border = {
  left: {
    style: BorderStyle.NIL,
    size: 0,
  },
  right: {
    style: BorderStyle.NIL,
    size: 0,
  },
  top: {
    style: BorderStyle.NIL,
    size: 0,
  },
  bottom: {
    style: BorderStyle.NIL,
    size: 0,
  },
};

const MTFborder = {
  left: {
    style: BorderStyle.SINGLE,
    size: 2,
  },
  right: {
    style: BorderStyle.SINGLE,
    size: 2,
  },
  top: {
    style: BorderStyle.SINGLE,
    size: 2,
  },
  bottom: {
    style: BorderStyle.SINGLE,
    size: 2,
  },
};
function headers() {
  return new Table({
    columnWidths: [6505, 6505],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [
              new Paragraph({
                text: "Student Name:.........................",
                bold: true,
                // heading: HeadingLevel.HEADING_1,
              }),
            ],
          }),
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            alignment: AlignmentType.RIGHT,
            children: [
              new Paragraph({
                text: "Roll Number:.............................",
                bold: true,
                // heading: HeadingLevel.HEADING_1,
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [
              new Paragraph({
                text: "Teacher's Name:......................",
                bold: true,
                // heading: HeadingLevel.HEADING_1,
              }),
            ],
          }),
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            alignment: AlignmentType.RIGHT,
            children: [
              new Paragraph({
                text: "Teacher's Sign:...........................",
                bold: true,
                // heading: HeadingLevel.HEADING_1,
              }),
            ],
          }),
        ],
      }),
    ],
  });
}
//
function optionsTabel(testimage) {
  // console.log(testimage)
  return new Table({
    columnWidths: [5505, 5505],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [
              displayOptions(testimage[0], testimage[4], testimage[5]),
            ],
          }),
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [
              displayOptions(testimage[1], testimage[4], testimage[5]),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [
              displayOptions(testimage[2], testimage[4], testimage[5]),
            ],
          }),
          new TableCell({
            borders: border,
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
            children: [
              displayOptions(testimage[3], testimage[4], testimage[5]),
            ],
          }),
        ],
      }),
    ],
  });
}

module.exports = {
  create,
  // buildDOCXwithCallback,
};
