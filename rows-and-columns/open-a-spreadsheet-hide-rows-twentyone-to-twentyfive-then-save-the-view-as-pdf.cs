// Title: Hide Rows 21‑25 in an Excel Sheet and Save as PDF using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to load (or create) an Excel workbook, conceal rows 21 through 25 on the first worksheet, and generate a PDF that respects the hidden rows via Aspose.Cells' SaveFormat.Pdf.
// Keywords: Aspose.Cells | C# PDF export | Excel hide rows | HideRows method | SaveFormat.Pdf | .NET spreadsheet to PDF | row visibility | Excel to PDF conversion | Aspose.Cells API | worksheet row hiding
// Common Searches: Aspose.Cells hide rows C# | Export Excel to PDF with hidden rows .NET | How to conceal specific rows before PDF conversion Aspose | HideRows 20 5 Aspose.Cells example | C# generate PDF from Excel while hiding rows
// Developer Intent: Conceal a range of rows in an Excel worksheet and produce a PDF that omits them.
// Use Cases: Producing client‑ready reports that exclude internal calculation rows. | Creating printable invoices where detailed line items are hidden. | Generating presentation PDFs that show only summary data. | Automating compliance documents that must mask confidential rows before distribution.
// AI Prompts: Generate C# code with Aspose.Cells to hide rows 30‑35 and export the sheet to PDF. | Explain zero‑based indexing for the HideRows method in Aspose.Cells. | Add comprehensive error handling for missing source files during Excel‑to‑PDF conversion with Aspose.Cells. | Show how to hide rows dynamically based on a condition before saving as PDF.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example shows how to load (or create) an Excel workbook, conceal rows 21 through 25 on the first worksheet, and generate a PDF that respects the hidden rows via Aspose.Cells' SaveFormat.Pdf.
    public class HideRowsAndSavePdf
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Ensure the input workbook exists; create a simple one if missing
            if (!File.Exists(inputPath))
            {
                Workbook tempWb = new Workbook();
                Worksheet ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath);
                Console.WriteLine($"Created placeholder workbook: {inputPath}");
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide rows 21 to 25 (zero‑based index: 20 to 24)
            worksheet.Cells.HideRows(20, 5);

            // Save the workbook as PDF, preserving hidden rows in the view
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved to: {outputPath}");
        }
    }
}
