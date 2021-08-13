var express = require("express");
const { buildPDFWithCallback } = require("../service/print/pdf");
const { buildCSVWithCallback } = require("../service/print/csv");
const { buildCSVReportWithCallback } = require('../service/print/csvreport');
const { buildDOCXWithCallback } = require('../service/print/docx')
const requestMiddleware = require("../middlewares/request.middleware");
// const base64 = require('base64topdf');
const BASE_URL = "/program/v1";
// Refactor this to move to service
async function printDocx(req,res){
  const id = req.query.id;
  const format = req.query.format;
  // buildDOCXwithCallback(function (binary, error, errorMsg) {
    buildDOCXWithCallback(id,function (binary, error, errorMsg) {
    // console.log("Enttere dres")
    var date = new Date();
    if (!error) {
      if (format === "json") {
        const resJSON = {
          id: "api.collection.print",
          ver: "1.0",
          ts: date.toISOString(),
          params: {
            id,
            format,
            status: "successful",
            err: null,
            errmsg: null,
          }, 
          responseCode: "OK",
          result: {
            content_id: id,
            base64string: binary,
          },
        };
        res.send(resJSON);
      } else {
        
        res.setHeader('Content-Disposition', 'attachment; filename=MyDocument.docx');
        res.send(Buffer.from(binary, 'base64'));
      }
    } else {
      res.status(404).send({
        error: errorMsg,
      });
    }
  });
}

async function printPDF(req, res) {
  const id = req.query.id;
  const format = req.query.format;
  buildPDFWithCallback(id, function (binary, error, errorMsg) {
    var date = new Date();
    if (!error) {
      if (format === "json") {
        const resJSON = {
          id: "api.collection.print",
          ver: "1.0",
          ts: date.toISOString(),
          params: {
            id,
            format,
            status: "successful",
            err: null,
            errmsg: null,
          },
          responseCode: "OK",
          result: {
            content_id: id,
            base64string: binary,
          },
        };
        res.send(resJSON);
      } else {
        res.setHeader('Content-disposition', `attachment; filename=${id}.pdf`);
        res.setHeader('Content-type', 'application/pdf');
        res.send(`data:application/pdf;base64, ${binary}`);
      }
    } else {
      res.status(404).send({
        error: errorMsg,
      });
    }
  });
}
async function printCSV(req, res) {
  const id = req.query.id;
  const format = req.query.format;
  if((req.body.path).includes('print/csv')){
    console.log("CSV",req.body.path);
     buildCSVWithCallback(id, function (binary, error, errorMsg,filename) {
      var date = new Date();
      if (!error) {
        if (format === "json") {
          const resJSON = {
            id: "api.collection.print",
            ver: "1.0",
            ts: date.toISOString(),
            params: {
              id,
              format,
              status: "successful",
              err: null,
              errmsg: null,
            },
            responseCode: "OK",
            result: {
              content_id: id,
              base64string: binary,
            },
          };
          res.send(resJSON);
        } else {
          res.setHeader('Content-disposition', `attachment; filename=${filename}.csv`);
          res.setHeader('Content-type', 'text/csv; charset=utf-8');
          res.send(binary);
        }
      } else {
        res.status(404).send({
          error: errorMsg,
        });
      }
    });
  } else {
   buildCSVReportWithCallback(id, function (binary, error, errorMsg,filename) {
      var date = new Date();
      if (!error) {
        if (format === "json") {
          const resJSON = {
            id: "api.collection.print",
            ver: "1.0",
            ts: date.toISOString(),
            params: {
              id,
              format,
              status: "successful",
              err: null,
              errmsg: null,
            },
            responseCode: "OK",
            result: {
              content_id: id,
              base64string: binary,
            },
          };
          res.send(resJSON);
        } else {
          res.setHeader('Content-disposition', `attachment; filename=${filename}.csv`);
          res.setHeader('Content-type', 'text/csv; charset=utf-8');
          res.send(binary);
        }
      } else {
        res.status(404).send({
          error: errorMsg,
        });
      }
    });
  }
}
module.exports = function (app) {
  app
    .route(BASE_URL + "/print/pdf")
    .get(
      requestMiddleware.gzipCompression(),
      requestMiddleware.createAndValidateRequestBody,
      printDocx
      // printPDF
    );
  app
    .route(BASE_URL + "/print/csv")
    .get(
      requestMiddleware.gzipCompression(),
      requestMiddleware.createAndValidateRequestBody,
      printCSV
    );
  app
    .route(BASE_URL + "/print/report/aggregate")
    .get(
      requestMiddleware.gzipCompression(),
      requestMiddleware.createAndValidateRequestBody,
      printCSV
    );
  app
    .route(BASE_URL + "/print/docx")
    .get(
      requestMiddleware.gzipCompression(),
      requestMiddleware.createAndValidateRequestBody,
      // printDocx
      printPDF
    );
};