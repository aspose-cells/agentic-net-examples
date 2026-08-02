// Title: Select worksheets beginning with “Report” via LINQ and apply Letter paper size – Aspose.Cells for .NET
// Description: Loads a workbook, uses a case‑insensitive LINQ query to find all worksheets whose names start with "Report", sets each sheet's PageSetup.PaperSize to Letter, creates the output folder if needed, and saves the modified file.
// Keywords: Aspose.Cells | LINQ worksheet filter | worksheet name prefix | custom paper size | PageSetup.PaperSize | C# .NET
// Common Searches: LINQ filter worksheets by prefix Aspose.Cells | Set Letter paper size for selected sheets C# | How to change page setup for multiple worksheets | Case‑insensitive worksheet name selection Aspose.Cells | Apply custom print settings to report tabs
// Developer Intent: Identify worksheets whose names start with "Report" and assign a Letter paper size to each.
// Use Cases: Standardize print layout for all report tabs before exporting to PDF | Automate page‑setup configuration in a nightly reporting job | Enforce a consistent paper size across dynamically generated report sheets
// AI Prompts: Generate C# code that uses Aspose.Cells to select worksheets with names starting with "Report" (case‑insensitive) and set their PageSetup.PaperSize to PaperLetter. | Modify the example to use A4 paper size while keeping the same worksheet filter. | Extend the LINQ query to also include worksheets containing the word "Summary" in addition to the "Report" prefix.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// Loads a workbook, uses a case‑insensitive LINQ query to find all worksheets whose names start with "Report", sets each sheet's PageSetup.PaperSize to Letter, creates the output folder if needed, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Find worksheets whose names start with "Report" (case‑insensitive)
            var reportSheets = workbook.Worksheets
                                       .Cast<Worksheet>()
                                       .Where(ws => ws.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                                       .ToList();

            // Apply a custom paper size (Letter) to each selected worksheet
            foreach (Worksheet sheet in reportSheets)
            {
                // Set the paper size to Letter (8.5 x 11 inches)
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
