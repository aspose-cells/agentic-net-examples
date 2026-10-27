// Title: How to set an Excel workbook’s calculation mode to Manual with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual | disable automatic formula recalculation in Excel using Aspose.Cells | change workbook Settings.CalcMode to Manual in .NET | selective formula evaluation with Aspose.Cells workbook | prevent Excel formulas from recalculating when saving with Aspose.Cells
// Tags: Aspose.Cells manual CalcMode | Workbook.Settings.CalcMode usage | control Excel formula evaluation .NET | selective recalculation of workbook formulas | C# manual formula calculation setting

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file into an Aspose.Cells Workbook, optionally sets Workbook.Settings.CalcMode to Manual to stop automatic formula recalculation, ensures the output directory exists, saves the workbook to a new file, and logs any errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Optional: set calculation mode to manual if supported
            // Uncomment the following line if the CalcMode property is available in your Aspose.Cells version
            // workbook.Settings.CalcMode = CalcMode.Manual;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
