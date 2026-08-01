// Title: Unhide Column B, Set Width to 50 pts, and Export Worksheet to PDF with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, unhides column B, sets its width to 50 points, and saves the worksheet as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | unhide column | set column width points | export to PDF | PdfSaveOptions | Excel to PDF conversion | column B width 50
// Common Searches: Aspose.Cells unhide column and export to PDF | C# set column width in points before PDF conversion | make hidden column visible in PDF using Aspose.Cells | change column width in points with Aspose.Cells .NET
// Developer Intent: Make column B visible, define its width in points, and generate a PDF from the worksheet.
// Use Cases: Create a printable PDF report where previously hidden columns must appear with a specific width. | Prepare a PDF version of a spreadsheet after adjusting column layout for presentation or distribution. | Automate batch conversion of Excel files to PDF while ensuring certain columns are displayed with defined point widths.
// AI Prompts: Show C# code using Aspose.Cells to unhide column C, set its width to 30 points, and save the workbook as a PDF. | Explain the difference between setting column width in points versus characters in Aspose.Cells before PDF export. | Provide a script that applies the same unhide and width settings to every worksheet in a workbook prior to PDF conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsUnhideColumnToPdf
{
    // Loads an Excel workbook, unhides column B, sets its width to 50 points, and saves the worksheet as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Unhide column B (zero‑based index 1) and set its width to 50 points
            worksheet.Cells.UnhideColumn(1, 50.0);

            // Prepare PDF save options (default options are sufficient for this task)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Column B unhidden with width 50 points and saved to PDF: {outputPath}");
        }
    }
}
