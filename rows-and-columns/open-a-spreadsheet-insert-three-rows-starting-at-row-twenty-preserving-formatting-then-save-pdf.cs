// Title: C# – Insert rows at row 20 with formatting and export to PDF using Aspose.Cells
// Description: Loads an Excel workbook, inserts three rows beginning at row 20 while copying the style from the preceding row, updates formula references, and saves the result directly as a PDF.
// Keywords: Aspose.Cells C# insert rows | preserve formatting Aspose.Cells | save workbook as PDF | InsertRows method | CopyFormatType SameAsAbove | UpdateReference option | Excel to PDF conversion | row insertion example | Aspose.Cells PDF export
// Common Searches: Aspose.Cells insert rows at specific index | keep cell style when adding rows Aspose.Cells | export modified Excel to PDF with Aspose.Cells C# | insert multiple rows and update formulas Aspose.Cells | copy format from above row Aspose.Cells
// Developer Intent: Add three rows at row 20 without losing existing styles and generate a PDF file.
// Use Cases: Generate a PDF report with extra spacing rows for section headers. | Add placeholder rows for future data before creating an invoice PDF. | Insert rows to accommodate new line items while preserving table formatting for printable documents.
// AI Prompts: Write C# code that inserts three rows at row 20, copies the format from the row above, updates references, and saves the workbook as a PDF using Aspose.Cells. | Demonstrate how the UpdateReference flag changes formula ranges after inserting rows with Aspose.Cells. | Show an example of converting an edited Excel sheet to PDF after inserting rows, including error handling.

using System;
using Aspose.Cells;

namespace AsposeCellsInsertRowsAndSavePdf
{
    // Loads an Excel workbook, inserts three rows beginning at row 20 while copying the style from the preceding row, updates formula references, and saves the result directly as a PDF.
    class Program
    {
        static void Main()
        {
            // Path to the existing spreadsheet (can be .xlsx, .xls, etc.)
            string inputFile = "input.xlsx";

            // Path for the output PDF file
            string outputPdf = "output.pdf";

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (you can change the index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare insert options to preserve formatting (copy format from the row above)
            InsertOptions insertOptions = new InsertOptions
            {
                CopyFormatType = CopyFormatType.SameAsAbove,
                UpdateReference = true
            };

            // Insert three rows starting at row 20 (zero‑based index 19)
            worksheet.Cells.InsertRows(19, 3, insertOptions);

            // Save the modified workbook as a PDF file
            workbook.Save(outputPdf, SaveFormat.Pdf);
        }
    }
}
