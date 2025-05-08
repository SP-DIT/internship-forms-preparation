const fs = require('fs');
const { PDFDocument } = require('pdf-lib');
const csv = require('csv-parser');
const { patchDocument, PatchType, TextRun } = require('docx');

async function modifyMonthlyAssessmentPdf(name, weekRange, pdfBytes) {
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const form = pdfDoc.getForm();

    // Set the name of the student
    const nameField = form.getTextField('Name');
    nameField.setText(name);

    // Set the week range
    const weekDropdown = form.getDropdown('Weeks');
    weekDropdown.select(weekRange);

    const pdfBytesModified = await pdfDoc.save();
    return pdfBytesModified;
}

async function createMonthlyAssessment() {
    const pdfPath = 'Bimonthly Assessments (eFORM)_YearLong_March2024 v2_1.pdf';
    const namesCsvPath = 'names.csv';
    const pdfBytes = fs.readFileSync(pdfPath);

    const names = [];
    fs.createReadStream(namesCsvPath)
        .pipe(csv())
        .on('data', (row) => {
            names.push(row.name);
        })
        .on('end', async () => {
            console.log(names);
            for (const name of names) {
                const weekRanges = ['Weeks 1-8', 'Weeks 9-16', 'Weeks 17-24', 'Weeks 25-32', 'Weeks 33-40'];
                for (let i = 0; i < weekRanges.length; i++) {
                    const weekRange = weekRanges[i];
                    const modifiedPdfBytes = await modifyMonthlyAssessmentPdf(name, weekRange, pdfBytes);
                    const outputFileName = `./output/${name}_${weekRange.replace(' ', '_')}.pdf`;
                    fs.writeFileSync(outputFileName, modifiedPdfBytes);
                }
            }
            console.log('PDFs generated successfully.');
        });
}

async function modifyOverallAssessmentPdf(name, pdfBytes) {
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const form = pdfDoc.getForm();

    // Set the name of the student
    const nameField = form.getTextField('Name');
    nameField.setText(name);

    const pdfBytesModified = await pdfDoc.save();
    return pdfBytesModified;
}

async function createOverallAssessment() {
    const pdfPath = 'Overall Assessment (eFORM) v2.pdf';
    const namesCsvPath = 'names.csv';
    const pdfBytes = fs.readFileSync(pdfPath);

    const names = [];
    fs.createReadStream(namesCsvPath)
        .pipe(csv())
        .on('data', (row) => {
            names.push(row.name);
        })
        .on('end', async () => {
            console.log(names);
            for (const name of names) {
                const modifiedPdfBytes = await modifyOverallAssessmentPdf(name, pdfBytes);
                const outputFileName = `./output/${name}_Overall_Assessment.pdf`;
                fs.writeFileSync(outputFileName, modifiedPdfBytes);
            }
            console.log('PDFs generated successfully.');
        });
}

async function modifyReflectionDocx(name, weekRange, from, to, content) {
    return patchDocument({
        data: content,
        outputType: 'nodebuffer',
        patches: {
            name: {
                type: PatchType.PARAGRAPH,
                children: [new TextRun(name)],
            },
            week: {
                type: PatchType.PARAGRAPH,
                children: [new TextRun(weekRange)],
            },
            from: {
                type: PatchType.PARAGRAPH,
                children: [new TextRun(from)],
            },
            to: {
                type: PatchType.PARAGRAPH,
                children: [new TextRun(to)],
            },
        },
    });
}

async function createReflection() {
    const docxPath = 'Student Reflection.docx';
    const namesCsvPath = 'names.csv';
    const content = fs.readFileSync(docxPath);

    const names = [];
    const weekRanges = ['5-12', '13-20', '21-28', '29-40'];
    const dates = [
        { from: '14 April 2025', to: '30 June 2025' },
        { from: '30 June 2025', to: '25 Aug 2025' },
        { from: '25 Aug 2025', to: '20 Oct 2025' },
        { from: '20 Oct 2025', to: '12 Jan 2026' },
    ];
    fs.createReadStream(namesCsvPath)
        .pipe(csv())
        .on('data', (row) => {
            names.push(row.name);
        })
        .on('end', async () => {
            console.log(names);
            for (const name of names) {
                for (let i = 0; i < weekRanges.length; i++) {
                    const weekRange = weekRanges[i];
                    const { from, to } = dates[i];
                    const modifiedDocxBytes = await modifyReflectionDocx(name, weekRange, from, to, content);
                    const outputFileName = `./output/${name}_Student_Reflection_Weeks_${weekRange}.docx`;
                    fs.writeFileSync(outputFileName, modifiedDocxBytes);
                }
            }
            console.log('DOCXs generated successfully.');
        });
}

async function modifyStudentReportDocx(name, studentId, companyName, submissionDate, content) {
    return patchDocument({
        data: content,
        outputType: 'nodebuffer',
        patches: {
            name: {
                type: PatchType.PARAGRAPH,
                children: [new TextRun(name)],
            },
            student_id: {
                type: PatchType.PARAGRAPH,
                children: [new TextRun(studentId)],
            },
            company_name: {
                type: PatchType.PARAGRAPH,
                children: [new TextRun(companyName)],
            },
            submission_date: {
                type: PatchType.PARAGRAPH,
                children: [new TextRun(submissionDate)],
            },
        },
    });
}

async function createStudentReport() {
    const docxPaths = ['FinalReport.docx', 'InterimReport.docx'];
    const submissionDates = ['9 Feb 2026', '15 Sep 2025'];
    const namesCsvPath = 'names.csv';

    const names = [];
    fs.createReadStream(namesCsvPath)
        .pipe(csv())
        .on('data', (row) => {
            names.push({ name: row.name, studentId: row.studentId, companyName: row.companyName });
        })
        .on('end', async () => {
            console.log(names);
            for (let i = 0; i < docxPaths.length; i++) {
                const docxPath = docxPaths[i];
                const submissionDate = submissionDates[i];
                const content = fs.readFileSync(docxPath);
                for (const { name, studentId, companyName } of names) {
                    const modifiedDocxBytes = await modifyStudentReportDocx(
                        name,
                        studentId,
                        companyName,
                        submissionDate,
                        content,
                    );
                    const outputFileName = `./output/${name}_${docxPath}`;
                    fs.writeFileSync(outputFileName, modifiedDocxBytes);
                }
                console.log('DOCXs generated successfully.');
            }
        });
}

createStudentReport().catch((err) => console.error(err));
