// Title: How to unhide rows 30‑35 in an Excel worksheet and save the workbook as PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unhide rows 30 to 35 before exporting to PDF | C# code to make hidden rows visible in an Excel file with Aspose.Cells | Save Excel workbook as PDF after revealing specific rows using Aspose.Cells .NET | Programmatically unhide multiple rows in an Aspose.Cells worksheet C# | Export Excel to PDF while handling previously hidden rows with Aspose.Cells
// Tags: unhide specific rows Aspose.Cells | row.IsHidden property C# Aspose.Cells | export workbook to PDF Aspose.Cells .NET | zero‑based row indexing Aspose.Cells | auto create rows when unhiding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input file, loads it into a Workbook, accesses the first worksheet, iterates rows 30‑35 (zero‑based indices 29‑34) and sets each row's IsHidden property to false (creating the row if it does not exist), ensures the output directory exists, and then saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing spreadsheet
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unhide rows 30 to 35 (Aspose.Cells uses zero‑based indexing)
            // Row 30 => index 29, Row 35 => index 34
            for (int i = 29; i <= 34; i++)
            {
                try
                {
                    // Access the row via the Cells collection; this creates the row if it does not exist
                    Row row = sheet.Cells.Rows[i];
                    row.IsHidden = false; // Unhide the row
                    // Height reset is optional; omitted to avoid compatibility issues
                }
                catch (Exception exRow)
                {
                    Console.WriteLine($"Failed to process row {i + 1}: {exRow.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
