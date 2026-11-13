// Title: Auto‑fit all rows on workbook load with Aspose.Cells for .NET to keep original layout
// Common Searches: C# Aspose.Cells auto fit rows after opening an existing Excel workbook | How to preserve row height formatting when loading an XLSX file using Aspose.Cells | Automatically adjust row heights for each worksheet on workbook load in .NET | Aspose.Cells iterate worksheets and call AutoFitRows on load | Load Excel file and auto‑size rows programmatically with Aspose.Cells C#
// Tags: Aspose.Cells AutoFitRows method | row height auto‑adjustment on workbook load | preserve Excel row height layout with Aspose.Cells | C# iterate worksheets for row sizing | load XLSX and apply row height scaling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input XLSX file, loads it into an Aspose.Cells Workbook, loops through each worksheet calling AutoFitRows() to automatically adjust row heights, and then saves the modified workbook, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Auto‑fit rows for each worksheet to adjust row heights
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.AutoFitRows();
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
