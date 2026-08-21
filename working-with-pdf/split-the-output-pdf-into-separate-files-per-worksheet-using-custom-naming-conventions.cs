// Title: C# – Split an Excel workbook into individual PDFs per worksheet with custom names using Aspose.Cells
// Description: Loads an Excel file, loops through its worksheets, and uses Aspose.Cells PdfSaveOptions.SheetSet to render each sheet as a separate PDF. The output files are automatically named after the corresponding worksheet, producing one PDF per sheet.
// Keywords: Aspose.Cells | C# | split workbook PDF | export worksheet to PDF | PdfSaveOptions SheetSet | custom PDF filename | Excel to PDF per sheet | batch PDF conversion | worksheet name file | .NET PDF generation
// Common Searches: export each Excel sheet to a separate PDF C# | Aspose.Cells split workbook into PDFs | how to name PDF files by worksheet name Aspose | PdfSaveOptions SheetSet example | C# generate PDF per worksheet | batch convert Excel sheets to PDFs
// Developer Intent: Produce a distinct PDF for every worksheet in an Excel workbook, using the sheet name as the PDF file name, with Aspose.Cells for .NET.
// Use Cases: Distribute departmental reports where each department's data resides on its own worksheet. | Automate creation of individual invoice PDFs when each invoice is stored on a separate sheet. | Share chart or dashboard sheets as standalone PDFs without exposing the full workbook. | Generate PDF packages for regulatory submissions, separating data sections by worksheet.
// AI Prompts: Write C# code with Aspose.Cells that converts each worksheet of an Excel file into a separate PDF named after the sheet. | Explain how PdfSaveOptions.SheetSet restricts PDF rendering to a single worksheet and how to apply it in a loop. | Provide a C# snippet that sanitizes worksheet names to create valid file names before saving PDFs. | Suggest performance optimizations for converting large workbooks with many sheets to PDFs using Aspose.Cells. | Create a PowerShell script that calls the compiled C# program to process multiple workbooks in a folder.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfSplitExample
{
    // Loads an Excel file, loops through its worksheets, and uses Aspose.Cells PdfSaveOptions.SheetSet to render each sheet as a separate PDF. The output files are automatically named after the corresponding worksheet, producing one PDF per sheet.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            string sourceFile = "input.xlsx";
            Workbook workbook = new Workbook(sourceFile);

            // Iterate through each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Create PDF save options and limit rendering to the current sheet only
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // SheetSet accepts an array of zero‑based sheet indexes
                    SheetSet = new SheetSet(new int[] { i })
                };

                // Build a custom file name using the worksheet name
                string outputFile = $"{sheet.Name}.pdf";

                // Save the single‑sheet PDF
                workbook.Save(outputFile, pdfOptions);
            }
        }
    }
}
