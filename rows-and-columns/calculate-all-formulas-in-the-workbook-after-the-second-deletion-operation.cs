// Title: Recalculate all workbook formulas after deleting the second row and second column using Aspose.Cells for .NET
// AI Prompts: Load an Excel file, delete row index 1, delete column index 1 with reference updates, recalculate all formulas, and save the result with Aspose.Cells in C#. | Using Aspose.Cells, remove the second row and second column from the first worksheet, trigger a full formula recalculation, and export the modified workbook. | In C#, apply DeleteRows and DeleteColumns (updateReference=true) on a workbook, then call CalculateFormula to refresh all calculations before saving.
// Common Searches: Aspose.Cells calculate formulas after deleting a column in C# | how to recalculate workbook after removing rows with Aspose.Cells .NET | delete second row then second column and refresh formulas using Aspose.Cells | C# Aspose.Cells DeleteRows DeleteColumns recalculate all formulas example
// Tags: Aspose.Cells DeleteRows API | Aspose.Cells DeleteColumns with reference update | Aspose.Cells CalculateFormula after structural changes | C# recalculate workbook formulas Aspose.Cells | Excel row and column deletion Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, deletes the second row and then the second column (updating references), recalculates every formula in the workbook, and saves the modified file to a new location using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // ----- First deletion operation -----
            // Delete the second row (zero‑based index 1)
            sheet.Cells.DeleteRows(1, 1);

            // ----- Second deletion operation -----
            // Delete the second column (zero‑based index 1) and update references
            sheet.Cells.DeleteColumns(1, 1, true);

            // Recalculate all formulas after deletions
            workbook.CalculateFormula();

            // Save the updated workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
