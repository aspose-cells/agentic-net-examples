// Title: Aspose.Cells C# – Insert Two Rows at Index 50 and Export Worksheet to PDF
// Description: Loads an existing Excel file, inserts two rows at zero‑based index 50 on the first worksheet while preserving the surrounding default formatting, and saves the result directly as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | InsertRows method | row index 50 | default formatting | SaveFormat.Pdf | Excel to PDF conversion | worksheet row insertion | programmatic Excel manipulation | Aspose.Cells API
// Common Searches: Aspose.Cells insert rows at row 50 C# | C# add rows to Excel and save as PDF | Insert multiple rows with default style using Aspose.Cells | Convert modified Excel to PDF with Aspose.Cells .NET | How to use InsertRows method in Aspose.Cells | Save Excel workbook as PDF after row insertion Aspose
// Developer Intent: Insert two rows at row index 50 in a worksheet and then save the workbook as a PDF.
// Use Cases: Add blank rows before a data block in a template, then generate a PDF report for stakeholders. | Programmatically create space for new entries in an existing spreadsheet and export the updated view as a printable PDF. | Modify an Excel file by inserting rows without altering formatting, and produce a PDF version for distribution.
// AI Prompts: Generate C# code with Aspose.Cells to insert three rows at row 20 and export the workbook to PDF. | Explain how Aspose.Cells preserves default formatting when inserting rows and how to customize the style afterward. | Provide robust error handling for missing input files and PDF write permissions while inserting rows and converting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsInsertRowsAndPdf
{
    // Loads an existing Excel file, inserts two rows at zero‑based index 50 on the first worksheet while preserving the surrounding default formatting, and saves the result directly as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (create/load rule)
            string inputFile = "input.xlsx";
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert two rows at index 50 (zero‑based). Default formatting (same as above) is applied.
            worksheet.Cells.InsertRows(50, 2);

            // Save the modified workbook as a PDF (save rule)
            string outputPdf = "output.pdf";
            workbook.Save(outputPdf, SaveFormat.Pdf);
        }
    }
}
