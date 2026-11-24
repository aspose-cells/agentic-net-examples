// Title: Use LINQ to filter worksheets whose names start with "Report" and set their paper size to Letter with Aspose.Cells for .NET
// Common Searches: how to use LINQ to select worksheets by name prefix in Aspose.Cells C# | set paper size to Letter for specific worksheets using Aspose.Cells | filter workbook worksheets case-insensitive with LINQ Aspose.Cells | apply standard paper size to multiple sheets in a .NET Excel file | Aspose.Cells change page setup for worksheets that start with Report
// Tags: LINQ worksheet filtering Aspose.Cells | set worksheet paper size Aspose.Cells | case-insensitive worksheet name match C# | apply predefined paper size Letter Aspose.Cells | select worksheets by prefix using Cast

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one if missing), uses LINQ to retrieve all worksheets whose names begin with "Report" (ignoring case), sets each selected sheet's PageSetup.PaperSize to PaperLetter, ensures the output directory exists, saves the workbook, and logs any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";   // path to the source workbook
        const string outputPath = "output.xlsx"; // path where the modified workbook will be saved

        try
        {
            // Load the workbook – if the file does not exist, create a new empty workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Select worksheets whose names start with "Report" (case‑insensitive)
            var reportSheets = workbook.Worksheets
                                       .Cast<Worksheet>()
                                       .Where(ws => ws.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase));

            // Apply a standard paper size to each selected worksheet
            foreach (var sheet in reportSheets)
            {
                // Use a predefined paper size (e.g., Letter) instead of a custom size
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
