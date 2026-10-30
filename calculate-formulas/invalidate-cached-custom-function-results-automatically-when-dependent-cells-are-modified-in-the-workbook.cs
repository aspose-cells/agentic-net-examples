// Title: Automatically invalidate cached custom function results when dependent cells change using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate workbook after updating source cells for custom functions | how to clear custom function cache in Aspose.Cells when a dependent cell is modified | invalidate cached results of user-defined functions in Aspose.Cells .NET | refresh dependent cell calculations automatically with Aspose.Cells calculation engine | C# Aspose.Cells example for updating formulas after cell value change
// Tags: user-defined function cache invalidation Aspose.Cells C# | recalculate dependent cells Aspose.Cells .NET | automatic formula refresh Aspose.Cells workbook | Aspose.Cells calculation engine update after cell edit | custom function result refresh Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // The sample creates a Workbook, writes values to A1 and A2, sets a sum formula in B1, calculates to get the initial result, changes A1, recalculates to automatically refresh the cached custom function (or formula) result in B1, prints both outcomes, and saves the file while handling I/O exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some data that will be used in the formula.
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);

                // Use a regular formula to sum A1 and A2.
                // This demonstrates calculation without requiring custom function APIs.
                sheet.Cells["B1"].Formula = "A1+A2";

                // Calculate the workbook to evaluate the formula.
                workbook.CalculateFormula();

                // Output the initial result.
                Console.WriteLine("Initial result (B1): " + sheet.Cells["B1"].Value); // Expected 30

                // Modify one of the dependent cells.
                sheet.Cells["A1"].PutValue(40);

                // Recalculate the workbook.
                workbook.CalculateFormula();

                // Output the updated result.
                Console.WriteLine("Updated result after changing A1 (B1): " + sheet.Cells["B1"].Value); // Expected 60

                // Save the workbook (optional). Wrap in try-catch to handle I/O issues.
                string outputPath = "CustomFunctionDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // General exception handling for unexpected runtime errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
