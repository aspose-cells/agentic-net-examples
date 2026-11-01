// Title: Export a specific worksheet (e.g., Sheet1) to CSV while removing leading empty rows and columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export selected worksheet to CSV without leading blank rows in C# | How to save only one sheet as CSV using TxtSaveOptions in .NET | C# remove empty top rows and columns when converting Excel to CSV with Aspose | Set active sheet before saving to CSV with Aspose.Cells | Create output folder automatically when exporting worksheet to CSV in C# | Trim leading blank rows and columns during CSV export using Aspose.Cells
// Tags: selected worksheet CSV export Aspose.Cells | TxtSaveOptions CSV conversion .NET | trim leading empty rows columns CSV Aspose | ensure output directory before workbook.Save | set active sheet index Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads 'input.xlsx', selects the worksheet named 'Sheet1', makes it the active sheet, configures TxtSaveOptions for CSV format, creates the destination folder if missing, and saves the active worksheet to 'output.csv', automatically omitting leading empty rows and columns.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the worksheet named "Sheet1".
            Worksheet sheet = workbook.Worksheets["Sheet1"];
            if (sheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Set the retrieved worksheet as the active sheet.
            workbook.Worksheets.ActiveSheetIndex = sheet.Index;

            // Configure CSV export options using TxtSaveOptions (compatible with all versions).
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
            // Example: set the delimiter explicitly (optional).
            // csvOptions.Separator = ',';

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the active worksheet to a CSV file.
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Worksheet exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
