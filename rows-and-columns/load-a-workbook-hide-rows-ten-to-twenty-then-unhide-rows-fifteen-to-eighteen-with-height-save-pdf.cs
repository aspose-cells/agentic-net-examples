// Title: Hide rows 10‑20, set custom height for rows 15‑18, and export the worksheet to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide rows 10 to 20 then unhide rows 15 to 18 with specific height before PDF export | C# code to set row.IsHidden false and row.Height for a range of rows using Aspose.Cells | Export Excel worksheet to PDF after changing row visibility with Aspose.Cells for .NET | How to adjust row height for selected rows in Aspose.Cells prior to saving as PDF | Hide and unhide rows in an Excel file using Aspose.Cells and then convert to PDF in C#
// Tags: hide rows Aspose.Cells C# | unhide rows set height Aspose.Cells | row visibility manipulation Aspose.Cells | export worksheet to PDF Aspose.Cells | modify row height before PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, hides rows 10‑20, unhides rows 15‑18 while setting their height to 20 points, and saves the worksheet as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Hide rows 10 to 20 (1‑based indexing, zero‑based in API)
            for (int i = 9; i <= 19; i++)
            {
                var row = worksheet.Cells.Rows[i];
                row.IsHidden = true;
            }

            // Unhide rows 15 to 18 and set a specific height
            for (int i = 14; i <= 17; i++)
            {
                var row = worksheet.Cells.Rows[i];
                row.IsHidden = false;
                row.Height = 20; // example height
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
