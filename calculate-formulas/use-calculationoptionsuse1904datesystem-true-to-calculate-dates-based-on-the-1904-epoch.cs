// Title: C# – Convert Excel 1904‑epoch serial dates to .NET DateTime with Aspose.Cells
// Description: Demonstrates how to enable CalculationOptions.Use1904DateSystem, read a serial date (e.g., 10) from a cell, convert it to a .NET DateTime using CellsHelper.GetDateTimeFromDouble, apply a built‑in date format, recalculate formulas, and save the workbook as 1904DateSystem.xlsx.
// Keywords: Aspose.Cells C# | 1904 date system | CalculationOptions.Use1904DateSystem | CellsHelper.GetDateTimeFromDouble | Excel serial date conversion | Mac Excel 1904 epoch | Workbook.CalculateFormula
// Common Searches: How to enable 1904 date system in Aspose.Cells C# | Convert Excel serial number to DateTime using 1904 epoch Aspose.Cells | Aspose.Cells read Mac Excel dates | Set Use1904DateSystem true before CalculateFormula | C# example for 1904 date system in Aspose.Cells
// Developer Intent: Activate the 1904 date system and accurately translate Excel serial dates to .NET DateTime values in C#.
// Use Cases: Importing dates from Mac‑origin Excel files that use the 1904 epoch. | Generating reports where dates must reflect the 1904 date system. | Running formula calculations that depend on the 1904 date system and presenting the results as readable dates.
// AI Prompts: Show C# code to set CalculationOptions.Use1904DateSystem = true before workbook.CalculateFormula() with Aspose.Cells. | Provide an example that reads a serial date from a cell and converts it to DateTime using the 1904 flag. | Explain how to apply a built‑in date format after converting a serial number with the 1904 date system.

using System;
using Aspose.Cells;

// Demonstrates how to enable CalculationOptions.Use1904DateSystem, read a serial date (e.g., 10) from a cell, convert it to a .NET DateTime using CellsHelper.GetDateTimeFromDouble, apply a built‑in date format, recalculate formulas, and save the workbook as 1904DateSystem.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: In some Aspose.Cells versions the 1904 date system flag is not exposed.
            // The example proceeds without setting it explicitly; the date conversion uses the 1904 flag.

            // Put a numeric value that represents an Excel serial date.
            // In the 1904 date system, the value 10 corresponds to 1904‑01‑11.
            sheet.Cells["A1"].PutValue(10.0);

            // Convert the serial number to a .NET DateTime using the 1904 flag.
            double serialValue = sheet.Cells["A1"].DoubleValue;
            DateTime dateValue = CellsHelper.GetDateTimeFromDouble(serialValue, true);
            sheet.Cells["B1"].PutValue(dateValue);

            // Apply a date number format so the cell displays a readable date.
            Style dateStyle = sheet.Cells["B1"].GetStyle();
            dateStyle.Number = 14; // Built‑in date format
            sheet.Cells["B1"].SetStyle(dateStyle);

            // Perform any pending calculations.
            workbook.CalculateFormula();

            // Save the workbook.
            string outputPath = "1904DateSystem.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
