// Title: Hide rows 50‑55 in an Excel worksheet and export the visible area to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide rows 50 to 55 before saving workbook as PDF | How to export only visible rows to PDF with Aspose.Cells .NET | C# code to hide specific rows in Excel and generate PDF using Aspose.Cells | Skip hidden rows when converting Excel to PDF with Aspose.Cells | Hide a range of rows in Aspose.Cells and create PDF output
// Tags: hide rows Aspose.Cells C# | Excel worksheet PDF export Aspose.Cells | skip hidden rows during PDF conversion Aspose.Cells | row visibility management Aspose.Cells | convert visible area to PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file with Aspose.Cells, hides rows 50 through 55 on the first worksheet by setting their IsHidden property, and then saves the workbook as a PDF that contains only the visible rows, handling file‑existence checks and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide rows 50 to 55 (zero‑based indices 49 to 54)
            for (int rowIndex = 49; rowIndex <= 54; rowIndex++)
            {
                // Aspose.Cells creates rows on demand, so this works even if the row is empty
                sheet.Cells.Rows[rowIndex].IsHidden = true;
            }

            // Export the visible part of the worksheet to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
