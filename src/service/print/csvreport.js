const { getData } = require('./csvImporter')
const JSON2CSV = require('json2csv').parse
const axios = require('axios')
const buildCSVReportWithCallback = async (id, callback) => {
  let error = false
  let errorMsg = ''
  let totalMarks = 0
  let report = []
  let projectIdList = await getQuestionFromProject(id)
  // console.log("Projectlist:", projectIdList.result.content);
  if (projectIdList.result.content) {
    let questionPaperContent = []
    for (const id of projectIdList.result.content) {
      let questionContent
      await getData(id.identifier)
        .then(async data => {
          if (data.error) {
            callback(null, data.error, data.errorMsg, null)
          } else {
            let totalNodes = 0
            let totalAcceptedNodes = 0
            if (data.paperData.childNodes) {
              totalNodes = data.paperData.childNodes.length
            }
            if (data.paperData.acceptedContents) {
              console.log('object')
              totalAcceptedNodes = data.paperData.acceptedContents.length
            }
            const examName = data.paperData.name
            questionContent = {
              QuestionPperCollection: examName,
              TotalQuestionCreated: totalNodes,
              TotalQuestionApproved: totalAcceptedNodes
            }
            questionPaperContent.push(questionContent)
          }
          // console.log(" Json:",questionPaperContent);
        })
        .catch(e => {
          console.log(e)
          error = true
          errorMsg = ''
          callback(null, error, errorMsg, null)
        })
    }
    // console.log(" Json:",questionPaperContent);
    let fields = [
          'QuestionPperCollection',
          'TotalQuestionCreated',
          'TotalQuestionApproved'
        ]
        // console.log('Final Json:', questionPaperContent)
        let csv = JSON2CSV(questionPaperContent, {
          fields: fields,
          withBOM: true
        })
        let filename = 'QuestionReport'
        filename = filename.replace(/\s/g, '')
        return callback(csv, error, errorMsg, filename)
  }
}
const getQuestionFromProject = async id => {
  const headers = {
    'Content-Type': 'application/json'
  }
  let data = {
    request: {
      filters: {
        programId: 'ea2035b0-e53c-11eb-862f-7fd867b01374',
        objectType: 'collection',
        status: ['Draft'],
        primaryCategory: ['Question Paper']
      }
    }
  }
  let url = `${envVariables.baseURL}/action/composite/v3/search`
  try {
    var result = await axios.post(url, data, { headers: headers })
    return result.data
  } catch (error) {
    var errData = { data: { msg: error.message } }
    return errData
  }
}
module.exports = {
  buildCSVReportWithCallback
}