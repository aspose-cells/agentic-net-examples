// Title: C# – Insert Two Blank Rows at Row 51 and Convert Excel to PDF with Aspose.Cells
// Description: Loads an existing .xlsx file, inserts two empty rows at the zero‑based index 50 (row 51) using Aspose.Cells’ InsertRows method with default formatting, then saves the workbook directly as a PDF.
// Keywords: Aspose.Cells InsertRows C# | worksheet.Cells.InsertRows | Excel to PDF conversion .NET | SaveFormat.Pdf Aspose.Cells | add blank rows Aspose.Cells | zero‑based row index Excel | C# Excel row insertion example
// Common Searches: Aspose.Cells insert rows at specific position C# | How to add blank rows before exporting to PDF with Aspose.Cells | Insert rows at row 51 using Aspose.Cells .NET | Convert modified Excel workbook to PDF in C# | Default formatting when inserting rows with Aspose.Cells
// Developer Intent: Add two empty rows at row 51 in a worksheet and immediately generate a PDF version of the workbook.
// Use Cases: Create space for new data entries in a financial report before producing a PDF for stakeholders. | Insert placeholder rows in a template, then export the finalized sheet as a printable PDF. | Add spacing in a schedule worksheet to improve readability in the generated PDF document.
// AI Prompts: Write C# code that inserts a configurable number of rows at a given zero‑based index with default formatting using Aspose.Cells and then saves the workbook as a PDF. | Show how to verify row insertion (e.g., row count or cell values) before converting the workbook to PDF with Aspose.Cells. | Provide an example that includes page‑setup options when exporting an Excel sheet with newly inserted rows to PDF.

using System;
using Aspose.Cells;

namespace AsposeCellsInsertRowsAndPdf
{
    // Loads an existing .xlsx file, inserts two empty rows at the zero‑based index 50 (row 51) using Aspose.Cells’ InsertRows method with default formatting, then saves the workbook directly as a PDF.
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file
            string inputFile = "input.xlsx";

            // Path for the generated PDF file
            string outputFile = "output.pdf";

            // Load the workbook from the file (uses the Workbook(string) constructor)
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (you can change the index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert two rows at index 50 (zero‑based). This uses the default InsertRows method,
            // which creates empty rows with default formatting.
            worksheet.Cells.InsertRows(50, 2);

            // Save the modified workbook as a PDF (uses the Save method with SaveFormat.Pdf)
            workbook.Save(outputFile, SaveFormat.Pdf);
        }
    }
}
