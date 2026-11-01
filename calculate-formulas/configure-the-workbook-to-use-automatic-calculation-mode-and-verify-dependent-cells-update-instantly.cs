// Title: Set Automatic Calculation Mode in Aspose.Cells for .NET and Verify Real‑Time Formula Updates
// Common Searches: C# Aspose.Cells how to enable automatic calculation for workbook formulas | Aspose.Cells .NET verify dependent cells recalculate after source cell change | example of automatic formula recalculation using Aspose.Cells workbook | Aspose.Cells set calculation mode to Automatic and test cell dependency | C# code to check real‑time update of formula cells in Aspose.Cells
// Tags: Aspose.Cells automatic calculation mode | C# formula dependency update Aspose.Cells | Aspose.Cells workbook calculation setting | real‑time formula recalculation Aspose.Cells | Aspose.Cells save workbook to XLSX

using System;
using System.IO;
using Aspose.Cells;

// The example creates an in‑memory workbook, assigns a value to A1, defines B1 with a formula that depends on A1, confirms the initial calculation, changes A1, reads B1 to demonstrate automatic recalculation, and finally saves the workbook as an XLSX file.
class AutomaticCalculationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // Set up a simple dependency:
            // A1 = 10
            // B1 = =A1 * 2   (formula dependent on A1)
            // -------------------------------------------------
            cells["A1"].PutValue(10);
            cells["B1"].Formula = "A1*2";

            // Force calculation to ensure initial values are computed
            workbook.CalculateFormula();

            // Verify initial dependent value
            Console.WriteLine($"Initial B1 value (should be 20): {cells["B1"].Value}");

            // -------------------------------------------------
            // Change the source cell (A1) and verify that B1 updates instantly
            // -------------------------------------------------
            cells["A1"].PutValue(15); // Update source value

            // Because the default calculation mode is Automatic,
            // accessing B1 triggers recalculation
            Console.WriteLine($"Updated B1 value (should be 30): {cells["B1"].Value}");

            // -------------------------------------------------
            // Save the workbook to a file (optional, demonstrates save rule)
            // -------------------------------------------------
            string outputPath = "AutomaticCalculationDemo.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
