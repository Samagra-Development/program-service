const { getData } = require('./csvImporter')
const JSON2CSV = require('json2csv').parse
var cheerio = require('cheerio')

const buildCSVWithCallback = async (id, callback) => {
  let error = false
  let errorMsg = ''
  let totalMarks = 0
  getData(id)
    .then(async data => {
      if (data.error) {
        callback(null, data.error, data.errorMsg, null)
      } else {
        let subject
        let grade
        if(data.paperData.subject && data.paperData.gradeLevel){
          subject = data.paperData.subject[0]
          grade = data.paperData.gradeLevel[0]
        }
        const examName = data.paperData.name

        data.sectionData.forEach(d => {
          d.questions.forEach((element, index) => {
            const marks = parseInt(d.section.children[index].marks)
            if (!isNaN(marks)) totalMarks += marks
          })
        })

        const questionPaperContent = []
        let questionCounter = 0
        for (const d of data.sectionData) {
          for (const [index, question] of d.questions.entries()) {
            questionCounter += 1
            let questionContent
            let blooms
            let learningOutcome
            let chaperName 

            if (question.category === 'MCQ') {
              if (question.learningOutcome && question.learningOutcome[0]) {
                learningOutcome = question.learningOutcome[0]
              } else {
                learningOutcome = ''
              }

              if (question.bloomsLevel && question.bloomsLevel[0]) {
                blooms = question.bloomsLevel[0]
              } else {
                blooms = ''
              }
              if (blooms === 'Remember') {
                blooms = 'Knowledge'
              } else if (blooms === 'Understand') {
                blooms = 'Understanding'
              } else if (blooms === 'Apply') {
                blooms = 'Application'
              }
              if (question.topic && question.topic[0]) {
                chaperName = question.topic[0]
              } else {
                chaperName = ''
              }
              questionContent = await renderMCQ(
                question,
                questionCounter,
                grade,
                subject,
                examName,
                learningOutcome,
                blooms,
                chaperName
              )
              questionPaperContent.push(questionContent)
            }
          }
        }

        let fields = [
          'Class',
          'Subject',
          'TopicName',
          'Questions',
          'Option1',
          'Option2',
          'Option3',
          'Option4',
          'CorrectAnswer(1/2/3/4)',
          'Competencies',
          'Skills',
          'QuestionImageUrl',
          'ChapterName'
        ]

        let csv = JSON2CSV(questionPaperContent, {
          fields: fields,
          withBOM: true
        })
        let filename = grade + '_' + subject + '_' + examName
        filename = filename.replace(/\s/g, '')
        callback(csv, error, errorMsg, filename)
      }
    })
    .catch(e => {
      console.log(e)
      error = true
      errorMsg = ''
      callback(null, error, errorMsg, null)
    })
}

const cleanHTML = (str, nbspAsLineBreak = false) => {
  // Remove HTML characters since we are not converting HTML to PDF.
  return str
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, nbspAsLineBreak ? '\n' : '')
}

function extractTextFromElement (elem) {
  let rollUp = ''
  if (cheerio.text(elem)) return cheerio.text(elem)
  else if (elem.name === 'sup')
    return { text: elem.children[0].data, sup: true }
  else if (elem.name === 'sub')
    return { text: elem.children[0].data, sub: true }
  else if (elem.type === 'text' && elem.data) return elem.data
  else {
    if (elem.children && elem.children.length) {
      for (const nestedElem of elem.children) {
        let recurse = extractTextFromElement(nestedElem)
        if (Array.isArray(rollUp)) {
          rollUp.push(recurse)
        } else {
          if (Array.isArray(recurse)) {
            rollUp = recurse
          } else if (typeof recurse === 'object') {
            rollUp = [rollUp, recurse]
          } else rollUp += recurse
        }
      }
    }
  }
  return rollUp
}

async function getStack (htmlString, questionCounter) {
  const stack = []
  let count = 0;
  let p = 0
  $ = cheerio.load(htmlString)
  const elems = $('body')
    .children()
    .toArray()
  for (const [index, elem] of elems.entries()) {
    let nextLine = ''
    switch (elem.name) {
      case 'p':
        let extractedText = extractTextFromElement(elem)
        // Returns array if superscript/subscript inside
        // console.log("para:",extractedText)
        if (Array.isArray(extractedText)) {
          nextLine = { text: extractedText }
        } else {
          nextLine += extractedText
        }
        nextLine = { text: nextLine }
        break
      case 'ol':
        nextLine = {
          ol: elem.children.map(
            el =>
              el.children[0] &&
              (el.children[0].data ||
                (el.children[0].children[0] && el.children[0].children[0].data))
          )
        }
        break
      case 'ul':
        nextLine = {
          ul: elem.children.map(
            el =>
              el.children[0] &&
              (el.children[0].data ||
                (el.children[0].children[0] && el.children[0].children[0].data))
          )
        }
        break
      case 'figure':
        if (count === 0) {
          let { style } = elem.attribs
          let width = 1
          if (style) {
            width = parseFloat(
              style
                .split(':')
                .pop()
                .slice(0, -2)
            )
            width = width / 100
          }

          if (elem.children && elem.children.length) {
            let { src } = elem.children[0].attribs
         
            if (src.startsWith('data:image/png')) {
              nextLine = ''
            } else if (src.startsWith('data:image/jpeg')){
              nextLine = ''
            }else{
              src = src.replace('/assets/public','')
              count++
              nextLine = `${envVariables.QUE_IMG_URL}` + src
            }
          }
          if (!nextLine)
            nextLine = '<An image of an unsupported format was scrubbed>'
        }
        break
    }
    if (index === 0 && questionCounter) {
      if (elem.name === 'p') {
        if (typeof nextLine === 'object')
          nextLine = { text: `${nextLine.text}` }
        else nextLine = `${nextLine}`
      } else stack.push(`${nextLine}`)
    }
    stack.push(nextLine)
  }
  return stack
}

async function renderMCQ (
  question,
  questionCounter,
  grade,
  subject,
  examName,
  learningOutcome,
  blooms,
  topic
) {
  
  const questionOptions = [],
    answerOptions = ['A', 'B', 'C', 'D']
  let questionTitle
  let finalQuestion = ''

  for (const [index, qo] of question.editorState.options.entries()) {
   
    let qoBody = qo.value.body
    let qoData =
      qoBody.search('img') >= 0 ||
      qoBody.search('sup') >= 0 ||
      qoBody.search('sub') >= 0 ||
      qoBody.match(/<p>/g).length > 1 
      // qoBody.match(/<ol>/g).length >= 1
        ? await getStack(qoBody, answerOptions[index])
        : [`${cleanHTML(qoBody)}`]
    questionOptions.push(qoData)
  }

  let q = question.editorState.question

  questionTitle =
    q.search('img') >= 0 ||
    q.search('sub') >= 0 ||
    q.search('sup') >= 0 ||
    q.match(/<p>/g).length > 1
      ? await getStack(q, questionCounter)
      : [`${cleanHTML(q)}`]

  let answer = ''
  for (const option of question.options) {
    if (option.answer === true) {
      answer = option.value.resindex + 1
    }
  }

  let imageurl = envVariables.QUE_IMG_URL
  let queurl = ''

  for (let que of questionTitle) {
    if (typeof que === 'object') {
      finalQuestion += que.text
        } else {
      if (que.includes(imageurl)) {
        queurl = que
    } else if(que.match('<An image of an unsupported format was scrubbed>')){
        queurl = que
     }else{
        finalQuestion = que
      }
    }
  }

  let data = {
    Class: grade,
    Subject: subject,
    TopicName: examName,
    Questions: finalQuestion,
    Option1: questionOptions[0][0],
    Option2: questionOptions[1][0],
    Option3: questionOptions[2][0],
    Option4: questionOptions[3][0],
    'CorrectAnswer(1/2/3/4)': answer,
    Competencies: learningOutcome,
    Skills: blooms,
    QuestionImageUrl: queurl,
    ChapterName: topic
  }
  return data
  
}

module.exports = {
  buildCSVWithCallback
}