// Title: Freezing the header row and first column while updating all formulas in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file, use Worksheet.FreezePanes to lock the header row and the A column, call Workbook.CalculateFormula to update every formula, then save the modified workbook with Aspose.Cells in C#. | In C#, apply FreezePanes at row index 1 and column index 1 on the first worksheet, trigger a full formula recalculation, and write the result to a new Excel file using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# freeze top row and left column then refresh formulas | How to recalculate formulas after applying FreezePanes with Aspose.Cells .NET | C# example for freezing panes and updating dependent cells in an existing Excel workbook using Aspose.Cells
// Tags: Worksheet.FreezePanes Aspose.Cells C# | Workbook.CalculateFormula Aspose.Cells | freeze first row and first column Excel Aspose | recalculate formulas after freezing panes .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample loads input.xlsx, freezes the header row and first column via Worksheet.FreezePanes(1,1,0,0), recalculates all formulas with Workbook.CalculateFormula, ensures the output directory exists, and saves the updated workbook to output.xlsx while handling missing files and exceptions.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze panes at row 2 and column 2 (freezes the first row and first column)
            // Use the 4‑parameter overload to ensure compatibility with the current API
            sheet.FreezePanes(1, 1, 0, 0);

            // Recalculate all formulas so that dependent cells reflect the frozen view
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
