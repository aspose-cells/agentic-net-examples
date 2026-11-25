// Title: Recalculate all formulas after freezing panes in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate formulas after applying FreezePanes | How to refresh dependent cells when freezing header row with Aspose.Cells | Calculate workbook formulas programmatically after setting freeze panes in .NET | C# Aspose.Cells update formulas after using FreezePanes on a worksheet | Re‑evaluate Excel formulas after FreezePanes with Aspose.Cells API
// Tags: FreezePanes formula recalculation Aspose.Cells | Workbook.CalculateFormula after freezing rows | C# recalculate dependent cells Aspose.Cells | Excel freeze panes evaluate formulas .NET | Aspose.Cells update totals after FreezePanes

using System;
using Aspose.Cells;

namespace FreezePanesRecalculateDemo
{
    // The example creates a new workbook, fills it with quantity and unit price data, sets total and grand‑total formulas, freezes the header row at A2, calls Workbook.CalculateFormula to recalculate all formulas after freezing, prints the evaluated values, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (default format is XLSX)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some sample data
                worksheet.Cells["A1"].PutValue("Quantity");
                worksheet.Cells["B1"].PutValue("Unit Price");
                worksheet.Cells["C1"].PutValue("Total");

                // Fill rows with data
                for (int i = 2; i <= 6; i++)
                {
                    worksheet.Cells[$"A{i}"].PutValue(i - 1);               // Quantity 1..5
                    worksheet.Cells[$"B{i}"].PutValue(10 * (i - 1));       // Unit Price 10,20,...
                }

                // Set formula for Total column (C) to multiply Quantity * Unit Price
                for (int i = 2; i <= 6; i++)
                {
                    worksheet.Cells[$"C{i}"].Formula = $"=A{i}*B{i}";
                }

                // Add a grand total formula at the bottom
                worksheet.Cells["C7"].Formula = "=SUM(C2:C6)";
                worksheet.Cells["B7"].PutValue("Grand Total:");

                // Freeze panes so that the header row stays visible
                // Freeze at cell A2 (row index 1, column index 0) with 1 frozen row and 0 frozen columns
                worksheet.FreezePanes(1, 0, 1, 0);
                // Alternative using cell name:
                // worksheet.FreezePanes("A2", 1, 0);

                // Recalculate all formulas after freezing panes
                workbook.CalculateFormula();

                // Verify that formulas have been evaluated
                Console.WriteLine("Calculated values after freezing panes:");
                for (int i = 2; i <= 6; i++)
                {
                    Console.WriteLine($"Row {i}: Quantity={worksheet.Cells[$"A{i}"].IntValue}, " +
                                      $"Unit Price={worksheet.Cells[$"B{i}"].DoubleValue}, " +
                                      $"Total={worksheet.Cells[$"C{i}"].DoubleValue}");
                }
                Console.WriteLine($"Grand Total (C7) = {worksheet.Cells["C7"].DoubleValue}");

                // Save the workbook to a file
                string outputPath = "FreezePanesRecalculateDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
