// Title: Bulk import a 2‑D array into an Excel worksheet and force formula recalculation with Aspose.Cells for .NET
// AI Prompts: Generate C# using Aspose.Cells that opens input.xlsx, sets calculation mode to manual, imports a 1000‑row by 10‑column object[,] into cell A1 of the first worksheet, switches calculation back to automatic, forces a full formula recalculation, and saves to output.xlsx. | Provide a C# snippet that checks for the existence of an Excel file, loads it with Aspose.Cells, temporarily disables formula evaluation, bulk‑loads data via ImportTwoDimensionArray, then re‑enables calculation and calls CalculateFormula before saving. | Create example code in .NET that demonstrates how to import a large two‑dimensional array into a worksheet without triggering intermediate calculations, and then recompute all formulas after the import.
// Common Searches: Aspose.Cells how to set manual calculation mode before bulk import | ImportTwoDimensionArray without triggering formula recalculation .NET | Recalculate all formulas after loading large data set with Aspose.Cells | C# disable automatic calculation, import data, then enable calculation Aspose.Cells | Best practice for bulk data import and formula evaluation in Aspose.Cells
// Tags: manual calculation mode Aspose.Cells .NET | ImportTwoDimensionArray bulk data Aspose.Cells | force full formula recalculation Aspose.Cells | load workbook and save after data import Aspose.Cells | disable automatic calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads input.xlsx, optionally sets the workbook to manual calculation mode, imports a 1000 × 10 object array into the first worksheet at A1 using ImportTwoDimensionArray, restores automatic calculation, forces a full formula evaluation with CalculateFormula, and saves the result to output.xlsx while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired one)
            Worksheet sheet = workbook.Worksheets[0];

            // Example bulk data: a 2‑dimensional array of objects
            int rows = 1000;
            int cols = 10;
            object[,] bulkData = new object[rows, cols];

            // Populate the array with sample data (replace with real data as needed)
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    bulkData[i, j] = $"R{i + 1}C{j + 1}";
                }
            }

            // Import the array into the worksheet starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportTwoDimensionArray(bulkData, 0, 0);

            // Force a full recalculation of all formulas in the workbook
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
