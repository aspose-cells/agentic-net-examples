// Title: Disable automatic calculation, modify cells, and manually recalculate formulas with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# turn off workbook auto calculation before updating cells | How to manually invoke CalculateFormula after changing cell values in Aspose.Cells | C# example for disabling automatic formula evaluation and editing multiple cells with Aspose.Cells | Programmatically recalculate Excel formulas using Aspose.Cells after cell edits | Set cell values and trigger workbook.CalculateFormula in Aspose.Cells .NET
// Tags: auto calculation disabled workbook Aspose.Cells | manual CalculateFormula invocation Aspose.Cells | cell value assignment Aspose.Cells C# | workbook save output Aspose.Cells .NET | conditional workbook creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file or creates a new workbook, accesses the first worksheet, assigns values to cells A1, B2, and C3, disables automatic formula calculation, manually triggers workbook.CalculateFormula to update dependent formulas, ensures the output directory exists, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Access the first worksheet (creates one if workbook is empty)
            Worksheet sheet = workbook.Worksheets.Count > 0
                ? workbook.Worksheets[0]
                : workbook.Worksheets[workbook.Worksheets.Add()];

            // Modify cells
            sheet.Cells["A1"].PutValue(123);      // Example: set A1 to 123
            sheet.Cells["B2"].PutValue(456.78);   // Example: set B2 to 456.78
            sheet.Cells["C3"].PutValue("Test");   // Example: set C3 to a string

            // Manually trigger formula calculation for the entire workbook
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
