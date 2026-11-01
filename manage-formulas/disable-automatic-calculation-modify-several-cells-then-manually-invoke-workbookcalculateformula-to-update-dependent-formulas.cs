// Title: Manually control formula calculation in Aspose.Cells for .NET: disable auto‑calc, update cells, and invoke CalculateFormula
// Common Searches: Aspose.Cells C# set calculation mode to manual before editing cells | Update cell values and recalculate dependent formulas using Workbook.CalculateFormula in .NET | Prevent automatic formula recalculation in Aspose.Cells and trigger manual calculation | How to modify cells and then manually recalculate all formulas with Aspose.Cells | Saving an Excel workbook after manual formula calculation using Aspose.Cells for .NET
// Tags: Aspose.Cells manual calculation configuration | update cell values programmatically Aspose.Cells | invoke manual formula evaluation Aspose.Cells | disable automatic formula recalculation Aspose.Cells | save Excel workbook after manual calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to (when supported) set Aspose.Cells to manual calculation mode, change values in cells A1 and A2, assign a formula to B1 that depends on those cells, manually recalculate all formulas with Workbook.CalculateFormula, and finally save the workbook as an Excel file while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // NOTE: In some Aspose.Cells versions the CalcMode property may not be available.
            // If needed, you can set the calculation mode to Manual as shown below (uncomment if supported):
            // workbook.Settings.CalcMode = CalculationMode.Manual;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Modify several cells
            cells["A1"].PutValue(10);      // Example input value
            cells["A2"].PutValue(20);      // Another input value
            cells["B1"].Formula = "=A1+A2"; // Dependent formula (will not recalc automatically if manual mode is set)

            // Manually calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "Result.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
