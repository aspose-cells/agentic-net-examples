// Title: How to refresh linked shapes after modifying source cells in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Update cell values and trigger a refresh of all linked shapes in a workbook with Aspose.Cells C#. | Recalculate formulas and refresh linked objects after changing source cells using the Aspose.Cells API. | Save the workbook ensuring linked pictures display the new data after cell updates in .NET.
// Common Searches: Aspose.Cells .NET refresh linked picture after cell value change | C# recalculate formulas and update linked shapes in Excel file | How to force linked objects to refresh in Aspose.Cells after editing cells | Save workbook with refreshed linked shapes using Aspose.Cells for .NET
// Tags: linked shape refresh Aspose.Cells | source cell update linked objects | formula recalculation C# Aspose.Cells | output directory handling Aspose.Cells | save workbook with refreshed linked pictures

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, changes the values of cells A1 and B2, recalculates all formulas to refresh any linked shapes or pictures, ensures the output directory exists, and saves the updated workbook to a new file.
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

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Update the source cells that linked shapes depend on
            worksheet.Cells["A1"].PutValue(123); // example value
            worksheet.Cells["B2"].PutValue(456); // example value

            // Recalculate formulas and refresh linked objects (including linked shapes)
            workbook.CalculateFormula();

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
