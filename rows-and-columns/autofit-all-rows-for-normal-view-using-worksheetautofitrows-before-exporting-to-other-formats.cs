// Title: C# example: Auto‑fit all rows in an Aspose.Cells worksheet with text wrapping before saving as PDF
// AI Prompts: Write C# code that creates a workbook, enables text wrapping for cells, calls Worksheet.AutoFitRows, and saves the file as a PDF using Aspose.Cells. | Show how to populate long strings in a worksheet, apply row auto‑fit, and export the result to PDF in a .NET application. | Provide a step‑by‑step C# snippet that demonstrates row height adjustment with Worksheet.AutoFitRows followed by PDF generation.
// Common Searches: asp.net how to auto fit rows in Aspose.Cells before exporting to PDF | c# Aspose.Cells auto fit row height with wrapped text example | Worksheet.AutoFitRows usage for PDF generation in .NET | adjust row heights automatically Aspose.Cells C# | export worksheet to PDF after auto‑fitting rows Aspose.Cells
// Tags: auto fit rows Aspose.Cells | Worksheet.AutoFitRows C# example | export worksheet to PDF Aspose.Cells | text wrapping row height Aspose.Cells | row height adjustment before PDF export

using System;
using System.IO;
using Aspose.Cells;

namespace AutoFitRowsExample
{
    // The sample creates a workbook, inserts long and multi‑line text, enables text wrapping, calls Worksheet.AutoFitRows to adjust all row heights, and saves the worksheet as a PDF on the desktop.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that will require row height adjustment
            sheet.Cells["A1"].PutValue("This is a long text that will demonstrate the AutoFitRows functionality. It should wrap and increase the row height accordingly.");
            sheet.Cells["A2"].PutValue("Another line with\nmultiple line breaks\nto test row auto‑fitting.");
            sheet.Cells["B1"].PutValue("Short text");
            sheet.Cells["B2"].PutValue("Medium length text");

            // Enable text wrapping so that rows need to expand vertically
            Style wrapStyle = sheet.Cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            sheet.Cells["A1"].SetStyle(wrapStyle);
            sheet.Cells["A2"].SetStyle(wrapStyle);

            // Auto‑fit all rows in the worksheet before exporting
            sheet.AutoFitRows();

            // Export the workbook to PDF (or any other format) after auto‑fitting
            string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AutoFitRowsResult.pdf");
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook saved with auto‑fitted rows to: {outputPath}");
        }
    }
}
