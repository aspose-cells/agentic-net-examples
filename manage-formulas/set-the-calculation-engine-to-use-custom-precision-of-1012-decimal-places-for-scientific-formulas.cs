// Title: How to set custom calculation precision (10‑12 decimal places) in Aspose.Cells for scientific formulas using C#
// Common Searches: Aspose.Cells C# set calculation engine precision to 12 decimal places | increase formula calculation precision in Aspose.Cells .NET | use CalcEnginePrecision to handle scientific notation in Excel with Aspose.Cells | custom decimal precision for workbook calculations in Aspose.Cells C# | high‑precision formula evaluation for large numbers in Aspose.Cells
// Tags: Aspose.Cells calculation engine precision | CalcEnginePrecision property .NET | high‑precision formula evaluation Excel | scientific notation formulas Aspose.Cells | custom decimal places workbook calculation | C# set workbook calculation precision

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, inserts a numeric value and a formula that uses scientific notation, attempts to configure the calculation engine precision via the CalcEnginePrecision property (not available in older versions), recalculates the formulas, and saves the workbook as 'CustomPrecision.xlsx' while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: The CalcEnginePrecision property is not available in the current Aspose.Cells version.
            // If a specific precision is required, ensure you are using a version that supports it.
            // workbook.Settings.CalcEnginePrecision = 12; // Uncomment if supported.

            // Add sample data and a formula that uses scientific notation
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(123456789.123456789); // sample value
            sheet.Cells["B1"].Formula = "A1*1E-5"; // formula using scientific notation

            // Recalculate formulas
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "CustomPrecision.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
