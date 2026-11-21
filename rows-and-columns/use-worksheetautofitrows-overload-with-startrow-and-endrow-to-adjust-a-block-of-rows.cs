// Title: Auto‑fit a specific block of rows with Worksheet.AutoFitRows(startRow, endRow) in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to auto fit rows 2 to 5 only | Worksheet.AutoFitRows overload startRow endRow example | adjust row height for a range of rows using Aspose.Cells .NET | auto fit rows based on cell content in a selected range Aspose.Cells | programmatically set Excel row heights for specific rows C# Aspose.Cells
// Tags: Worksheet.AutoFitRows row range overload | auto fit rows Aspose.Cells .NET | adjust Excel row height programmatically C# | Aspose.Cells row height auto fit specific rows | set row height based on content Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, writes varying length text into cells A2‑A5, then calls Worksheet.AutoFitRows with startRow=1 and endRow=4 to automatically adjust the heights of rows 2‑5, and saves the file as AutoFitRowsResult.xlsx.
class AutoFitRowsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in rows 2‑5 (zero‑based indexing)
            sheet.Cells["A2"].PutValue("Short");
            sheet.Cells["A3"].PutValue("This is a longer text that may require more row height");
            sheet.Cells["A4"].PutValue("Medium length text");
            sheet.Cells["A5"].PutValue("Another very long piece of text that should cause the row to expand automatically");

            // Auto‑fit rows from index 1 (row 2) to index 4 (row 5)
            int startRow = 1; // zero‑based
            int endRow = 4;   // zero‑based

            // Apply auto‑fit to the specified row range
            sheet.AutoFitRows(startRow, endRow);

            // Define output file path
            string outputPath = "AutoFitRowsResult.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
