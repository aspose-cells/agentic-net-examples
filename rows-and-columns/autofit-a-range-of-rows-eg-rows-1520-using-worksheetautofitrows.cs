// Title: Auto‑fit rows 15‑20 in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to auto‑fit a specific range of rows in Excel | How to use Worksheet.AutoFitRows to adjust rows 15 to 20 in .NET | Programmatically set row height for multiple rows using Aspose.Cells | Auto‑fit rows in Excel workbook with Aspose.Cells without specifying options
// Tags: Aspose.Cells auto-fit rows range | Worksheet.AutoFitRows C# example | adjust Excel row height programmatically .NET | auto fit multiple rows Aspose.Cells | set row height automatically in Excel using Aspose

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, and calls Worksheet.AutoFitRows(14, 5) to automatically adjust the height of rows 15 through 20. It ensures the output directory exists, saves the file as "AutoFitRows.xlsx", and reports success while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Auto‑fit rows 15‑20 (zero‑based index: 14‑19)
            // Overload without options to avoid enum compatibility issues
            sheet.AutoFitRows(14, 5);

            // Define output file path
            string outputPath = "AutoFitRows.xlsx";

            // Ensure the directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
