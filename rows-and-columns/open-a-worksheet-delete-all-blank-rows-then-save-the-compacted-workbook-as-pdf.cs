// Title: Aspose.Cells C# – Remove Blank Rows and Export Excel to PDF
// Description: Load an Excel workbook with Aspose.Cells, call Worksheet.Cells.DeleteBlankRows() to purge empty rows, set PdfSaveOptions.PrintingPageType = IgnoreBlank to skip blank pages, and save the compact workbook as a PDF. Ideal for cleaning user‑generated sheets before distribution.
// Keywords: Aspose.Cells | C# | DeleteBlankRows | remove empty rows | Excel to PDF | PdfSaveOptions | IgnoreBlank | compact PDF | batch Excel processing | export worksheet as PDF
// Common Searches: Aspose.Cells delete blank rows C# | How to remove empty rows from Excel using Aspose.Cells | Save Excel as PDF without blank pages Aspose.Cells | C# code to compact Excel workbook before PDF export | Worksheet.Cells.DeleteBlankRows example
// Developer Intent: Strip all completely empty rows from a worksheet and generate a PDF that contains only populated data.
// Use Cases: Sanitize data‑entry forms before creating printable reports. | Compress large financial workbooks into lean PDFs for email distribution. | Automate nightly batch conversion of Excel logs to PDF archives without gaps. | Prepare regulatory filings where blank rows cause pagination errors.
// AI Prompts: Write C# code that iterates through every worksheet in a workbook, deletes blank rows, and saves each sheet as an individual PDF using Aspose.Cells. | Explain the impact of PdfSaveOptions.PrintingPageType.IgnoreBlank on page count after rows are removed. | Create a PowerShell script that calls a compiled .NET assembly to process a folder of .xlsx files, removing blank rows and exporting PDFs. | Design a unit test that confirms DeleteBlankRows does not affect rows containing formulas that return empty strings.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an Excel workbook with Aspose.Cells, call Worksheet.Cells.DeleteBlankRows() to purge empty rows, set PdfSaveOptions.PrintingPageType = IgnoreBlank to skip blank pages, and save the compact workbook as a PDF. Ideal for cleaning user‑generated sheets before distribution.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific worksheet as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Delete all blank rows in the worksheet
            worksheet.Cells.DeleteBlankRows();

            // Optional: configure PDF save options (e.g., ignore completely blank pages)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PrintingPageType = PrintingPageType.IgnoreBlank
            };

            // Save the compacted workbook as a PDF file
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            // Clean up resources
            workbook.Dispose();

            Console.WriteLine("Workbook processed and saved as PDF successfully.");
        }
    }
}
