import React, { useState } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  HeadingLevel,
  AlignmentType,
  PageOrientation,
  TextRun,
} from "docx";

/**
 * Splits "Cat. No. and Course Name" into two parts: [catNo, courseName].
 */
function splitCatNoAndCourse(value) {
  const strVal = String(value || "").trim();
  const match = /^(\d+)\s+(.*)$/.exec(strVal);
  if (match) {
    return [match[1], match[2]];
  }
  return ["", strVal];
}

/**
 * Converts date-like strings into "Mon-YY Sem I/II" if possible; otherwise returns raw string.
 */
function formatSession(value) {
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return String(value || "");
  }
  const monthAbbr = date.toLocaleString("en-US", { month: "short" });
  const twoDigitYear = date.getFullYear().toString().slice(-2);
  const semester = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"].includes(
    monthAbbr
  )
    ? "Sem II"
    : "Sem I";
  return `${monthAbbr}-${twoDigitYear} ${semester}`;
}

/**
 * Returns single-line borders for a table cell.
 */
function allBorders() {
  return {
    top: { style: BorderStyle.SINGLE, size: 1 },
    bottom: { style: BorderStyle.SINGLE, size: 1 },
    left: { style: BorderStyle.SINGLE, size: 1 },
    right: { style: BorderStyle.SINGLE, size: 1 },
  };
}

/**
 * Creates multiple Paragraphs in one table cell—one paragraph per item in 'lines'.
 * Applies a custom style (TableData) to each paragraph so we can have a special font size.
 */
function multiLineCell(lines) {
  // 'lines' is an array of strings; each item gets its own paragraph in the cell
  const paragraphs = lines.map((text) => {
    return new Paragraph({
      style: "TableData", // Use the custom style for table data
      children: [new TextRun(String(text))],
    });
  });

  return new TableCell({
    children: paragraphs,
    borders: allBorders(),
  });
}

/**
 * Determines the academic year: "Aug-18 Sem I" => year=18; "Jan-18 Sem II" => year=17.
 */
function getAcademicYear(sessionLabel) {
  const semIRegex = /(Aug|Sep|Oct|Nov|Dec)-(\d{2,4})\s+Sem\s*I/i;
  const semIIRegex = /(Jan|Feb|Mar|Apr|May|Jun)-(\d{2,4})\s+Sem\s*II/i;

  let match = semIRegex.exec(sessionLabel);
  if (match) {
    return { year: match[2], matched: "SemI" };
  }

  match = semIIRegex.exec(sessionLabel);
  if (match) {
    // Jan => belongs to previous year
    const adjustedYear = String(parseInt(match[2], 10) - 1);
    return { year: adjustedYear, matched: "SemII" };
  }

  return { year: "Unknown", matched: null };
}

/**
 * Extracts column lines from an array of records => array of strings, one for each record.
 */
function getColumnLines(records, columnKey) {
  return records.map((r) => String(r[columnKey] || "").trim()).filter(Boolean);
}

const TranscriptGenerator = () => {
  const [selectedFile, setSelectedFile] = useState(null);

  const handleFileChange = (e) => {
    setSelectedFile(e.target.files[0] || null);
  };

  const handleGenerateTranscript = async () => {
    if (!selectedFile) {
      alert("Please select an Excel file first.");
      return;
    }

    try {
      // 1) Load the Excel file
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await selectedFile.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);

      const worksheet = workbook.worksheets[0];

      // 2) Extract rows from Excel
      const rowsData = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber >= 3) {
          const rowValues = row.values;
          rowsData.push({
            sem: rowValues[1] || "",
            catAndCourse: rowValues[2] || "",
            grade: rowValues[3] || "",
            semHrs: rowValues[4] || "",
          });
        }
      });

      // 3) Format each row
      const processed = rowsData.map((item) => {
        const formattedSem = formatSession(item.sem);
        const [catNo, courseName] = splitCatNoAndCourse(item.catAndCourse);
        return {
          rawSem: item.sem,
          Sem: formattedSem,
          Course: courseName,
          CatNo: catNo,
          Grade: item.grade,
          SemHrs: item.semHrs,
        };
      });

      // 4) Group rows by session-block until a blank line
      const groupedData = {};
      let currentSessionKey = null;

      processed.forEach((row) => {
        const rawVal = String(row.rawSem || "").trim();
        const courseVal = String(row.Course || "").trim();
        const gradeVal = String(row.Grade || "").trim();
        const hrsVal = String(row.SemHrs || "").trim();

        const isTrulyBlankRow = !rawVal && !courseVal && !gradeVal && !hrsVal;

        if (isTrulyBlankRow) {
          currentSessionKey = null;
          return;
        }

        if (rawVal) {
          currentSessionKey = row.Sem.trim();
        }
        if (!currentSessionKey) {
          currentSessionKey = "NoSession";
        }

        if (!groupedData[currentSessionKey]) {
          groupedData[currentSessionKey] = [];
        }

        // Skip lines that start with known "bottom" text
        const courseLower = row.Course.toLowerCase().trim();
        if (
          !courseLower.startsWith("total number of semester hours earned") &&
          !courseLower.startsWith(
            "number of semester hours required for graduation"
          ) &&
          !courseLower.startsWith("cumulative grade point average")
        ) {
          groupedData[currentSessionKey].push({
            Course: row.Course,
            CatNo: row.CatNo,
            Grade: row.Grade,
            SemHrs: row.SemHrs,
          });
        }
      });

      // 5) Build academicYears
      const academicYears = {};
      Object.entries(groupedData).forEach(([sessionLabel, courses]) => {
        const { year } = getAcademicYear(sessionLabel);
        if (!academicYears[year]) academicYears[year] = {};
        academicYears[year][sessionLabel] = courses;
      });

      // 6) Create doc
      const doc = new Document({
        creator: "Transcript Generator",
        description: "Generated transcript",
        title: "Generated Transcript",
        orientation: PageOrientation.PORTRAIT,
        sections: [
          {
            // Here is where "size" and "margins" typically go
            size: {
              // width: 12240, // 8.5 inches in Twips
              // height: 15840, // 11 inches in Twips
              orientation: PageOrientation.PORTRAIT,
            },
            // margins: {
            //   top: 1440, // 1 inch
            //   bottom: 1440, // 1 inch
            //   left: 1800, // 1.25 inches
            //   right: 1800, // 1.25 inches
            // },
            children: [
              // paragraphs, tables, etc.
            ],
          },
        ],
        styles: {
          default: {
            heading1: {
              run: {
                font: "Arial",
                size: 16, // 10pt
              },
            },
            heading2: {
              run: {
                font: "Arial",
                size: 16, // 10pt
              },
            },
            paragraph: {
              run: {
                font: "Arial",
                size: 16, // 10pt
              },
              spacing: {
                after: 100,
              },
            },
            document: {
              run: {
                font: "Arial",
                size: 16, // 10pt
              },
            },
            section: {
              run: {
                font: "Arial",
                size: 16, // 10pt
              },
            },
          },
          paragraphStyles: [
            {
              id: "TableData",
              name: "TableData",
              basedOn: "Normal",
              run: {
                font: "Arial",
                size: 16, // 12pt
              },
            },
          ],
        },
      });

      // Build paragraphs for the top part
      const firstGroup = [
        new Paragraph({
          text: "AFRICAN BIBLE COLLEGE",
          alignment: AlignmentType.CENTER,
          size: 24,
        }),
        new Paragraph({
          text: "P.O. BOX 1028, LILONGWE, MALAWI",
          alignment: AlignmentType.CENTER,
          size: 24,
        }),
        new Paragraph({
          text: "PHONE (265) 761-646 Email: registrar@abcmalawi.org",
          alignment: AlignmentType.CENTER,
          size: 24,
        }),
        new Paragraph(""), // blank line
      ];

      const secondGroup = [
        new Paragraph({
          text: "OFFICE OF THE REGISTRAR",
          alignment: AlignmentType.LEFT,
          size: 24,
        }),
        new Paragraph({
          text: "OFFICIAL TRANSCRIPT OF THE RECORD OF:",
          alignment: AlignmentType.LEFT,
          size: 24,
        }),
        new Paragraph({
          text: " [last name], [first name] STUDENT #[student number]",
          alignment: AlignmentType.LEFT,
          size: 24,
        }),
        new Paragraph({
          text: "BIRTHDATE: [mm-dd-yy]",
          alignment: AlignmentType.LEFT,
          size: 24,
        }),
        new Paragraph({
          text: "ATTENDANCE FROM: August 20[XX] TO: June 20[XX]",
          alignment: AlignmentType.LEFT,
          size: 24,
        }),
        new Paragraph({
          text: "PRESENT STATUS: GRADUATED [start year] WITH A BACHELORS OF [end year]",
          alignment: AlignmentType.LEFT,
          size: 24,
        }),
        new Paragraph({
          text: "CREDITS EARNED AT AFRICAN BIBLE COLLEGE",
          alignment: AlignmentType.LEFT,
          size: 24,
        }),
      ];

      // Combine paragraphs + table in one section
      const childrenArray = [...firstGroup, ...secondGroup];

      // 7) Build the table rows
      const rows = [];

      // Table header
      rows.push(
        new TableRow({
          children: [
            multiLineCell(["SESSION"]),
            multiLineCell(["COURSE"]),
            multiLineCell(["CAT. NO."]),
            multiLineCell(["GRADE"]),
            multiLineCell(["SEM. HRS."]),
            multiLineCell(["SESSION"]),
            multiLineCell(["COURSE"]),
            multiLineCell(["CAT. NO."]),
            multiLineCell(["GRADE"]),
            multiLineCell(["SEM. HRS."]),
          ],
        })
      );

      const allYearKeys = Object.keys(academicYears).filter(
        (y) => y !== "Unknown"
      );
      const numericYears = allYearKeys
        .map((y) => parseInt(y, 10))
        .filter((n) => !isNaN(n));
      numericYears.sort((a, b) => a - b);
      const sortedYears = numericYears.map(String);

      const leftoverYearKeys = allYearKeys.filter((y) =>
        isNaN(parseInt(y, 10))
      );
      leftoverYearKeys.forEach((y) => sortedYears.push(y));

      for (const yearStr of sortedYears) {
        const sessionsObj = academicYears[yearStr];
        const sessionLabels = Object.keys(sessionsObj);

        const leftRegex = /(Aug|Sep|Oct|Nov|Dec)-(\d{2,4})\s+Sem\s*I/i;
        const rightRegex = /(Jan|Feb|Mar|Apr|May|Jun)-(\d{2,4})\s+Sem\s*II/i;

        const leftKeys = sessionLabels.filter((s) => leftRegex.test(s));
        const rightKeys = sessionLabels.filter((s) => rightRegex.test(s));

        while (leftKeys.length > 0 || rightKeys.length > 0) {
          const leftKey = leftKeys.shift() || null;
          const rightKey = rightKeys.shift() || null;

          const leftCourses = leftKey ? sessionsObj[leftKey] : [];
          const rightCourses = rightKey ? sessionsObj[rightKey] : [];

          const leftCourseLines = getColumnLines(leftCourses, "Course");
          const leftCatNoLines = getColumnLines(leftCourses, "CatNo");
          const leftGradeLines = getColumnLines(leftCourses, "Grade");
          const leftHrsLines = getColumnLines(leftCourses, "SemHrs");

          const rightCourseLines = getColumnLines(rightCourses, "Course");
          const rightCatNoLines = getColumnLines(rightCourses, "CatNo");
          const rightGradeLines = getColumnLines(rightCourses, "Grade");
          const rightHrsLines = getColumnLines(rightCourses, "SemHrs");

          rows.push(
            new TableRow({
              children: [
                multiLineCell([leftKey || ""]),
                multiLineCell(leftCourseLines),
                multiLineCell(leftCatNoLines),
                multiLineCell(leftGradeLines),
                multiLineCell(leftHrsLines),
                multiLineCell([rightKey || ""]),
                multiLineCell(rightCourseLines),
                multiLineCell(rightCatNoLines),
                multiLineCell(rightGradeLines),
                multiLineCell(rightHrsLines),
              ],
            })
          );

          // blank row after pairing
          rows.push(
            new TableRow({
              children: Array.from({ length: 10 }, () => multiLineCell([""])),
            })
          );
        }
      }

      // Build the table
      const transcriptTable = new Table({
        rows,
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        columnWidths: [
          835, // SESSION (Left)
          2059, // COURSE (Left)
          533, // CAT. NO. (Left)
          547, // GRADE (Left)
          562, // SEM. HRS. (Left)
          835, // SESSION (Right)
          2059, // COURSE (Right)
          533, // CAT. NO. (Right)
          547, // GRADE (Right)
          562, // SEM. HRS. (Right)
        ],
      });

      // Insert the table right after secondGroup
      childrenArray.push(transcriptTable);

      // Then we add the final paragraphs right after the table
      childrenArray.push(new Paragraph(""));
      childrenArray.push(
        new Paragraph("Total Number of Semester Hours Earned: ")
      );
      childrenArray.push(
        new Paragraph("Number of Semester Hours Required for Graduation: ")
      );
      childrenArray.push(new Paragraph("Cumulative Grade Point Average: "));
      childrenArray.push(new Paragraph(""));
      childrenArray.push(
        new Paragraph(
          "The year consists of two semesters of approximately 16 weeks each.  Length of School Hour: Each lecture hour consists of not less than 50 minutes.  Grading System: A [100-96]; A- [95-93]; B+ [92-90]; B [89-87]; B- [86-84]; C+ [83-81]; C [80-78]; C- [77-75]; D+ [74-73]; D [72-71]; D- [70-66]; F [65- Below]; I [Incomplete}; W [Withdrew]."
        )
      );
      childrenArray.push(new Paragraph(""));
      childrenArray.push(new Paragraph(""));
      childrenArray.push(new Paragraph(""));
      childrenArray.push(new Paragraph("…………………………………………………………"));
      childrenArray.push(new Paragraph("ASSISTANT REGISTRAR"));
      childrenArray.push(
        new Paragraph(
          "This transcript is not valid unless it bears the seal of African Bible College."
        )
      );
      childrenArray.push(
        new Paragraph(`This transcript was issued on:  ${new Date()}.`)
      );

      // Finally, add everything in one single section
      doc.addSection({
        children: childrenArray,
      });

      // 8) Generate & download
      const blob = await Packer.toBlob(doc);
      saveAs(blob, "Generated_Transcript_Updated.docx");
      alert(
        "And the boyfriend of the year award goes to.....! Your transcript was successfully downloaded."
      );
    } catch (error) {
      console.error("Error generating transcript:", error);
      alert(
        "Error generating transcript. Please ensure the Excel document has headers in the following format: 'Sem. Cat.No. And Course Name	Grade	Sem.Hrs	GPA'."
      );
    }
  };

  return (
    <div style={{ margin: "20px" }}>
      <h3>Transcript Generator</h3>
      <input type="file" accept=".xls,.xlsx" onChange={handleFileChange} />
      <button onClick={handleGenerateTranscript} style={{ marginLeft: "10px" }}>
        Generate Transcript
      </button>
    </div>
  );
};

export default TranscriptGenerator;
