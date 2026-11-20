// Title: How to set Automatic calculation mode and confirm instant formula updates in an Aspose.Cells .NET workbook (C#)
// Common Searches: Aspose.Cells C# set workbook calculation mode to Automatic | verify dependent cell updates immediately after source cell change using Aspose.Cells | example of automatic formula recalculation with Aspose.Cells for .NET | force initial calculation then rely on automatic updates in Aspose.Cells C# | Aspose.Cells workbook automatic calculation mode demo code | check cell value after modifying referenced cell with Aspose.Cells C#
// Tags: set workbook calculation mode automatic Aspose.Cells | verify instant formula update Aspose.Cells | C# calculate formulas programmatically Aspose.Cells | dependent cell recalculation after source change .NET | save workbook with automatic calculation C#

using System;
using Aspose.Cells;

// The example creates a new Workbook, assigns a value to A1, defines B1 with a formula referencing A1, calls CalculateFormula for the initial evaluation, then changes A1 and shows that B1 updates instantly because the workbook operates in Automatic calculation mode, and finally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // The default calculation mode is Automatic, so no explicit setting is required.

            // Example: A1 = 10
            cells["A1"].PutValue(10);

            // B1 = A1 * 2 (dependent cell)
            cells["B1"].Formula = "A1*2";

            // Force calculation to ensure initial values are computed
            workbook.CalculateFormula();

            // Verify initial dependent value
            Console.WriteLine($"Initial B1 value (should be 20): {cells["B1"].Value}");

            // Change the source cell value
            cells["A1"].PutValue(15);

            // Because calculation mode is Automatic, B1 updates instantly
            Console.WriteLine($"Updated B1 value (should be 30): {cells["B1"].Value}");

            // Save the workbook (optional)
            string outputPath = "AutomaticCalculationDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
