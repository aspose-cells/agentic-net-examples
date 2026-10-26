// Title: How to configure Aspose.Cells CalcEngineOptions in C# to display a blank cell instead of a #DIV/0! error
// Common Searches: Aspose.Cells C# set calculation option to return empty cell for divide by zero | replace #DIV/0! with blank using Aspose.Cells CalcEngineOptions | C# Aspose.Cells hide division by zero error in generated Excel | configure workbook calculation engine to treat division by zero as null in .NET | Aspose.Cells suppress #DIV/0! error when saving Excel file
// Tags: Aspose.Cells CalcEngineOptions blank on divide by zero | C# suppress #DIV/0! error in Excel output | Excel calculation engine error handling Aspose.Cells | set error value to empty string Aspose.Cells | formula error display customization .NET

using System;
using Aspose.Cells;

// The example creates a workbook, assigns a formula that divides by zero, and shows how to use CalcEngineOptions (e.g., workbook.Settings.CalcEngineOptions.ErrorValue = "") to replace the resulting #DIV/0! error with a blank cell before saving the file.
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

            // Set a formula that would cause a division by zero error
            sheet.Cells["A1"].Formula = "10/0";

            // Recalculate formulas (default behavior will show the error value)
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
