// Title: Aspose.Cells .NET: Unhide All Rows in the First Worksheet (Default Height) and Export to PDF
// Description: Loads an Excel workbook, accesses the first worksheet, determines the allocated row count, unhides every row while preserving the default (auto‑fit) height, and saves the workbook as a PDF file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unhide rows .NET | export Excel to PDF C# | default row height Aspose | unhide hidden rows PDF conversion | Aspose.Cells workbook to PDF
// Common Searches: how to unhide rows with Aspose.Cells before PDF export | Aspose.Cells C# unhide all rows default height | convert Excel to PDF after showing hidden rows | C# code to unhide rows and save as PDF using Aspose
// Developer Intent: Reveal every row in the first worksheet using the default height and generate a PDF version of the workbook.
// Use Cases: Prepare a printable PDF from an Excel template that contains hidden rows. | Automate batch processing of workbooks to ensure all data is visible in the resulting PDFs. | Create a utility that receives an .xlsx path, unhides rows in the first sheet, and outputs a distribution‑ready PDF.
// AI Prompts: Generate a C# method that takes an input .xlsx file and an output PDF path, unhides all rows in the first worksheet with default height using Aspose.Cells, and saves the PDF. | Show sample Aspose.Cells code to unhide rows from row 0 to the last allocated row (height -1) and then export the workbook to PDF.

using System;
using Aspose.Cells;

namespace AsposeCellsUnhideRowsAndExportPdf
{
    // Loads an Excel workbook, accesses the first worksheet, determines the allocated row count, unhides every row while preserving the default (auto‑fit) height, and saves the workbook as a PDF file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the total number of rows in the sheet
            // Rows.Count returns the number of rows currently allocated in the sheet
            int totalRows = sheet.Cells.Rows.Count;

            // Unhide all rows in the first sheet.
            // The height parameter is set to -1 to keep the default (auto‑fit) height.
            sheet.Cells.UnhideRows(0, totalRows, -1);

            // Export the workbook to PDF (replace with your desired output path)
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
