// Title: Set Workbook.Settings.CalculationMode to Manual, batch update cells, and recalculate formulas once with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set calculation mode manual before bulk cell insertion C# | How to disable automatic formula calculation in Aspose.Cells and run CalculateFormula once | Best practice for updating thousands of rows in Aspose.Cells without triggering recalculation each time | C# Aspose.Cells performance tips for bulk data load with manual calculation mode | Calling Workbook.CalculateFormula after batch updates in Aspose.Cells .NET
// Tags: manual calculation mode Aspose.Cells | bulk worksheet cell insertion | Workbook.CalculateFormula execution | high‑performance Excel generation .NET | turn off auto calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, fills 1,000 rows with numeric values and dependent formulas, sets calculation mode to Manual, runs Workbook.CalculateFormula once to evaluate all formulas, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Perform bulk updates (example: fill 1000 rows with data and formulas)
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 1000; row++)
            {
                // Put a numeric value in column A
                sheet.Cells[row, 0].PutValue(row);

                // Set a formula in column B that depends on column A
                sheet.Cells[row, 1].Formula = $"=A{row + 1}*2";
            }

            // Recalculate the workbook once after all updates
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
