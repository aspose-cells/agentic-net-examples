// Title: Testing manual calculation mode so formulas stay unevaluated until Workbook.CalculateFormula is called in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells manual calculation mode example in C# | prevent automatic formula evaluation until CalculateFormula is invoked Aspose.Cells | verify cell value is null before calling Workbook.CalculateFormula in .NET | how to test manual calculation mode with Aspose.Cells workbook | C# code to control formula evaluation using Aspose.Cells
// Tags: manual calculation mode Aspose.Cells | Workbook.CalculateFormula method C# | disable automatic formula evaluation Aspose.Cells | formula caching verification .NET | testing formula evaluation control Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook, inserts a SUM formula into cell A1, confirms the cell's Value is null before invoking workbook.CalculateFormula(), then calls CalculateFormula() and verifies the cell's Value updates to the evaluated result (3).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Insert a formula into cell A1
            var cell = sheet.Cells["A1"];
            cell.Formula = "=SUM(1,2)";

            // Before calculation, the cell's cached value is null
            Console.WriteLine("Before CalculateFormula:");
            Console.WriteLine($"Cell A1 Value: {(cell.Value == null ? "null" : cell.Value.ToString())}");

            // Manually trigger formula calculation
            workbook.CalculateFormula();

            // After calculation, the cell should contain the evaluated result (3)
            Console.WriteLine("After CalculateFormula:");
            Console.WriteLine($"Cell A1 Value: {cell.Value}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
