// Title: Use LINQ to filter worksheets whose names start with "Report" and set Letter paper size with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses LINQ to locate all worksheets whose Name begins with "Report" and set each worksheet's PageSetup.PaperSize to PaperLetter via Aspose.Cells. | Adapt the example to apply an A4 paper size to the filtered worksheets instead of the default Letter size. | Add logic that logs a warning when no worksheets match the "Report" prefix before saving the workbook.
// Common Searches: aspnet linq filter worksheets by name prefix Aspose.Cells | set page setup paper size for multiple sheets Aspose.Cells .NET | how to apply Letter paper size to selected worksheets using Aspose.Cells | filter Excel worksheets starting with Report using C# and Aspose.Cells | apply page setup to worksheets with specific name pattern Aspose.Cells
// Tags: LINQ worksheet selection Aspose.Cells | set paper size page setup Aspose.Cells | filter worksheets by name prefix C# | apply Letter paper size Aspose.Cells | batch update worksheet page setup Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Linq;

// Loads an Excel workbook, uses LINQ to select worksheets whose names start with "Report", sets each selected worksheet's PageSetup.PaperSize to Letter, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Select worksheets whose names start with "Report"
            var reportSheets = workbook.Worksheets
                                       .Cast<Worksheet>()
                                       .Where(ws => ws.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase));

            // Apply a standard Letter paper size (8.5\" x 11\") to each selected worksheet
            foreach (var sheet in reportSheets)
            {
                var pageSetup = sheet.PageSetup;
                pageSetup.PaperSize = PaperSizeType.PaperLetter; // 8.5 x 11 inches
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
