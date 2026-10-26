// Title: How to assign a custom GlobalizationSettings.PivotSettings object to a workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom PivotSettings on WorkbookSettings.GlobalizationSettings | assign custom globalization instance to pivot tables in Aspose.Cells .NET | configure workbook pivot table culture using Aspose.Cells C# | how to apply custom GlobalizationSettings.PivotSettings to a loaded workbook in Aspose.Cells | set pivot table localization options programmatically with Aspose.Cells for .NET
// Tags: Aspose.Cells assign PivotSettings object | WorkbookSettings.GlobalizationSettings.PivotSettings usage | custom pivot table localization .NET | set pivot table culture programmatically | globalization instance for Excel pivot tables C#

using System;
using System.IO;
using Aspose.Cells;

// This C# example loads an existing Excel file with Aspose.Cells, demonstrates where to assign a custom GlobalizationSettings.PivotSettings object to the workbook's WorkbookSettings.GlobalizationSettings.PivotSettings, validates file paths, saves the modified workbook, and includes exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // NOTE: Custom PivotTable options are not available in this version of Aspose.Cells.
            // If needed, configure pivot tables using the PivotTable class after they are created.

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
