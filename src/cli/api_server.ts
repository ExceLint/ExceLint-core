//import express from 'express';
const express = require('express');
//import multer from 'multer';
const multer = require('multer');

const path = require('path');
import {ExcelJSON} from '../exceljson';
import {Colorize} from '../colorize';
import {WorkbookAnalysis} from '../ExceLintTypes';
import {Config} from '../config';
import {AnnotationData} from './bugs';

const app = express();
const port = 3003;
const upload = multer({dest: 'uploads/'});

const BUG_DATA_PATH = 'test/annotations-processed.json';
const theBugs = new AnnotationData(BUG_DATA_PATH);

app.route('/upload')
    .get((_req, res) => {
        // Return the HTML form for file upload
        res.send(`
      <!DOCTYPE html>
      <html>
      <body>
      <h2>Upload Excel File</h2>
      <form action="/upload" method="post" enctype="multipart/form-data">
        Select Excel file to upload:
        <input type="file" name="xlsfile" id="xlsfile">
        <input type="submit" value="Upload File">
      </form>
      </body>
      </html>
    `);
    })
    .post(upload.single('xlsfile'), (req, res) => {
        if (!req.file) {
            return res.status(400).send('No file uploaded.');
        }

        const file = req.file; // Uploaded file info
        const base = ''; // Adjust base as needed, possibly based on a request parameter

        // Initialize configurations and outputs array
        const outputs: WorkbookAnalysis[] = [];

        Config.setFormattingDiscount(50);
        Config.setReportingThreshold(0);

        const fname = path.join(file.destination, file.filename);

        console.warn('processing ' + fname);
        const inp = ExcelJSON.processWorkbook(base, fname);

        const facts = theBugs.check(inp);

        const output = Colorize.process_workbook(inp, '');
        outputs.push(output);


        // Respond with JSON
        res.json(outputs);
    });

app.listen(port, () => {
    console.log(`Server started at http://localhost:${port}`);
});
