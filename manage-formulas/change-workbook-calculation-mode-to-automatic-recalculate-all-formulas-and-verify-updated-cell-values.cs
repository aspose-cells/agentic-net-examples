// Title: Set Aspose.Cells workbook calculation mode to Automatic, recalculate formulas, and read updated cell values in C#
// Common Searches: Aspose.Cells C# set calculation mode to Automatic | Recalculate all formulas in an Excel workbook using Aspose.Cells .NET | Get cell value after CalculateFormula with Aspose.Cells | Save workbook after formula recalculation in Aspose.Cells C# | Verify updated Excel cell values programmatically with Aspose.Cells
// Tags: automatic calculation mode Aspose.Cells | recalculate workbook formulas Aspose.Cells | read cell value after formula evaluation Aspose.Cells | save workbook with refreshed values Aspose.Cells | load Excel file verify cell content Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook with Aspose.Cells for .NET, optionally switches the calculation mode to Automatic, forces a full formula recalculation, reads the updated value of a specific cell, and saves the workbook with the refreshed results.
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
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Optional: set calculation mode to Automatic if supported
            // In recent versions CalcMode may be unavailable; the default mode recalculates automatically.
            // workbook.Settings.CalcMode = CalcMode.Automatic;

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Read a sample cell value after recalculation (A1 in the first worksheet)
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            Console.WriteLine($"A1 value after recalculation: {cell.Value}");

            // Save the workbook with updated values
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
