// Title: How to unhide rows 30‑35 in an Excel worksheet and export the workbook to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unhide rows 30 to 35 before saving as PDF | C# code to make hidden rows visible in an Excel file with Aspose.Cells | Export Excel to PDF after revealing specific rows using Aspose.Cells .NET | Set row.IsHidden false for a range of rows in Aspose.Cells and generate PDF | Aspose.Cells .NET hide/unhide rows and convert workbook to PDF example
// Tags: unhide specific rows Aspose.Cells | row.IsHidden Aspose.Cells API | save workbook as PDF Aspose.Cells | Excel to PDF conversion Aspose.Cells .NET | default row height handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input file, loads it with Aspose.Cells, accesses the first worksheet, iterates rows 30‑35 (indices 29‑34) to set each row's IsHidden property to false, and then saves the workbook as a PDF file.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing spreadsheet
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unhide rows 30 to 35 (row numbers are 1‑based, indices are 0‑based)
            for (int rowIndex = 29; rowIndex <= 34; rowIndex++)
            {
                // Access the row via Cells.Rows collection
                Row row = sheet.Cells.Rows[rowIndex];
                row.IsHidden = false; // make the row visible
                // Height reset is optional; omitted to avoid API compatibility issues
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
