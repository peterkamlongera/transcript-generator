import React, { useState } from "react";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx"; // <-- For .xls -> .xlsx conversion in memory
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
 */
function multiLineCell(lines) {
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
 * Determines the academic year from a session label, e.g. "Aug-18 Sem I" => year=18; "Jan-18 Sem II" => year=17.
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
    // For "Jan-18 Sem II", the academic year is previous year: 17
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
      // 1) Read the file as ArrayBuffer
      const arrayBuffer = await selectedFile.arrayBuffer();

      // 2) If it's .xls, convert to .xlsx in memory using SheetJS
      const fileName = selectedFile.name.toLowerCase();
      let finalArrayBuffer = arrayBuffer; // Assume it's already .xlsx

      if (fileName.endsWith(".xls")) {
        // Convert .xls -> .xlsx
        const data = new Uint8Array(arrayBuffer);
        const sheetJSWorkbook = XLSX.read(data, { type: "array" });

        // Write it out as .xlsx array buffer
        const xlsxData = XLSX.write(sheetJSWorkbook, {
          bookType: "xlsx",
          type: "array",
        });
        finalArrayBuffer = xlsxData;
      }

      // 3) Parse .xlsx data with ExcelJS
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(finalArrayBuffer);

      // 4) Get first worksheet
      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        throw new Error("No worksheet found at index 0");
      }

      // 5) Extract rows (like your original code)
      //    We'll read from row 3 onward (if your data starts there):
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

      // 6) Format each row
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

      // 7) Group rows by session until a blank line
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

        // Skip lines that start with "bottom" text
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

      // 8) Build academic years
      const academicYears = {};
      Object.entries(groupedData).forEach(([sessionLabel, courses]) => {
        const { year } = getAcademicYear(sessionLabel);
        if (!academicYears[year]) academicYears[year] = {};
        academicYears[year][sessionLabel] = courses;
      });

      // 9) Create the docx
      const doc = new Document({
        creator: "Transcript Generator",
        description: "Generated transcript",
        title: "Generated Transcript",
        orientation: PageOrientation.PORTRAIT,
        sections: [
          {
            size: {
              orientation: PageOrientation.PORTRAIT,
            },
            children: [],
          },
        ],
        styles: {
          default: {
            heading1: {
              run: {
                font: "Arial",
                size: 16,
              },
            },
            heading2: {
              run: {
                font: "Arial",
                size: 16,
              },
            },
            paragraph: {
              run: {
                font: "Arial",
                size: 16,
              },
              spacing: {
                after: 100,
              },
            },
            document: {
              run: {
                font: "Arial",
                size: 16,
              },
            },
            section: {
              run: {
                font: "Arial",
                size: 16,
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
                size: 16,
              },
            },
          ],
        },
      });

      // Top paragraphs
      const firstGroup = [
        new Paragraph({
          text: "AFRICAN BIBLE COLLEGE",
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
          text: "P.O. BOX 1028, LILONGWE, MALAWI",
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
          text: "PHONE (265) 761-646 Email: registrar@abcmalawi.org",
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph(""), // blank line
      ];

      const secondGroup = [
        new Paragraph({
          text: "OFFICE OF THE REGISTRAR",
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          text: "OFFICIAL TRANSCRIPT OF THE RECORD OF:",
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          text: " [last name], [first name] STUDENT #[student number]",
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          text: "BIRTHDATE: [mm-dd-yy]",
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          text: "ATTENDANCE FROM: August 20[XX] TO: June 20[XX]",
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          text: "PRESENT STATUS: GRADUATED [start year] WITH A BACHELORS OF [end year]",
          alignment: AlignmentType.LEFT,
        }),
        new Paragraph({
          text: "CREDITS EARNED AT AFRICAN BIBLE COLLEGE",
          alignment: AlignmentType.LEFT,
        }),
      ];

      const childrenArray = [...firstGroup, ...secondGroup];

      // Build table rows
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

      // Sort academic years
      const allYearKeys = Object.keys(academicYears).filter(
        (y) => y !== "Unknown"
      );
      const numericYears = allYearKeys
        .map((y) => parseInt(y, 10))
        .filter((n) => !isNaN(n));
      numericYears.sort((a, b) => a - b);
      const sortedYears = numericYears.map(String);

      // leftover year keys
      const leftoverYearKeys = allYearKeys.filter((y) =>
        isNaN(parseInt(y, 10))
      );
      leftoverYearKeys.forEach((y) => sortedYears.push(y));

      // For each year, pair Sem I vs Sem II
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

          // Blank row after each pair
          rows.push(
            new TableRow({
              children: Array.from({ length: 10 }, () => multiLineCell([""])),
            })
          );
        }
      }

      const transcriptTable = new Table({
        rows,
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        columnWidths: [835, 2059, 533, 547, 562, 835, 2059, 533, 547, 562],
      });

      childrenArray.push(transcriptTable);

      // Final paragraphs
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
          "The year consists of two semesters of approximately 16 weeks each. Length of School Hour: Each lecture hour consists of not less than 50 minutes. Grading System: A [100-96]; A- [95-93]; B+ [92-90]; B [89-87]; B- [86-84]; C+ [83-81]; C [80-78]; C- [77-75]; D+ [74-73]; D [72-71]; D- [70-66]; F [65- Below]; I [Incomplete]; W [Withdrew]."
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

      // Add everything to the doc
      doc.addSection({
        children: childrenArray,
      });

      // 10) Generate & download
      const blob = await Packer.toBlob(doc);
      saveAs(blob, "Generated_Transcript.docx");
      alert(
        "Transcript successfully downloaded! And the boyfriend of the year award goes to...."
      );
    } catch (error) {
      console.error("Error generating transcript:", error);
      alert(`Error generating transcript: ${error.message}`);
    }
  };

  return (
    <div
      style={{
        backgroundColor: "#f4f4f4",
        padding: "20px",
        borderRadius: "10px",
        boxShadow: "0px 4px 10px rgba(0, 0, 0, 0.1)",
        maxWidth: "600px",
        margin: "20px auto",
      }}
    >
      <h3 style={{ textAlign: "center", color: "#333" }}>
        Transcript Generator
      </h3>
      <p style={{ textAlign: "center", color: "#555" }}>
        Upload an Excel file (.xlsx) to generate a transcript.
      </p>
      <input
        type="file"
        accept=".xls,.xlsx"
        onChange={handleFileChange}
        style={{
          display: "block",
          margin: "10px auto",
          padding: "10px",
          borderRadius: "5px",
          border: "1px solid #ccc",
        }}
      />
      <button
        onClick={handleGenerateTranscript}
        style={{
          display: "block",
          margin: "10px auto",
          padding: "10px 20px",
          backgroundColor: "#007BFF",
          color: "#fff",
          border: "none",
          borderRadius: "5px",
          cursor: "pointer",
        }}
      >
        Generate Transcript
      </button>
    </div>
  );
};

export default TranscriptGenerator;
