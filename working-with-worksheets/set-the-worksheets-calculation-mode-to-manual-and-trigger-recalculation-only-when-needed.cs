// Title: Set worksheet calculation mode to manual and recalculate formulas on demand with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set worksheet calculation mode to manual C# | Recalculate formulas only when needed using Aspose.Cells .NET | Disable automatic formula calculation in Aspose.Cells workbook | Trigger manual calculation after updating cell values with Aspose.Cells | Save workbook after setting manual calculation mode in Aspose.Cells C#
// Tags: manual calculation mode Aspose.Cells | explicit formula calculation C# | turn off automatic calculation Aspose.Cells | on-demand formula recalculation .NET | set worksheet calculation mode manual

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes a value to A1, assigns a formula to B1, sets the worksheet to manual calculation mode, explicitly recalculates formulas, ensures the output directory exists, and saves the file as ManualCalc.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Example modification that affects formulas
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);               // Input value
            sheet.Cells["B1"].Formula = "=A1*2";          // Formula dependent on A1

            // Recalculate formulas explicitly
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "ManualCalc.xlsx";

            // Ensure the directory exists (in case a relative path with directories is used)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
