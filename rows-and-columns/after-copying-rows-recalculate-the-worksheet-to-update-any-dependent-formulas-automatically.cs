// Title: Copy a worksheet row and recalculate dependent formulas using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy a row and refresh formulas | How to duplicate a row and recalculate worksheet formulas with Aspose.Cells | C# Aspose.Cells CalculateFormula after manual row copy | Using CalculationOptions Recursive true in Aspose.Cells to update dependent formulas | Copy row with Cell.Copy and recalculate formulas in Aspose.Cells .NET
// Tags: copy worksheet row Aspose.Cells C# | recalculate dependent formulas Aspose.Cells | worksheet.CalculateFormula recursive option | cell.Copy method row duplication | update formulas after row copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace RowCopyAndRecalculateDemo
{
    // The example creates a workbook, fills it with sample data and formulas, manually copies a source row to a new location using Cell.Copy, then forces a full recalculation of the worksheet with CalculationOptions (Recursive = true) to update all dependent formulas, and finally saves the result.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the output workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RowCopyRecalcDemo.xlsx");

                // Create a new workbook
                using (Workbook workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cells cells = worksheet.Cells;

                    // Populate sample data and formulas
                    cells["A1"].PutValue("Item");
                    cells["B1"].PutValue("Quantity");
                    cells["C1"].PutValue("Price");
                    cells["D1"].PutValue("Total");

                    cells["A2"].PutValue("Apple");
                    cells["B2"].PutValue(10);
                    cells["C2"].PutValue(0.5);
                    cells["D2"].Formula = "=B2*C2";

                    cells["A3"].PutValue("Banana");
                    cells["B3"].PutValue(5);
                    cells["C3"].PutValue(0.3);
                    cells["D3"].Formula = "=B3*C3";

                    // Save the workbook before copying (optional, for inspection)
                    string beforeCopyPath = Path.Combine(Directory.GetCurrentDirectory(), "BeforeCopy.xlsx");
                    workbook.Save(beforeCopyPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved before copying rows: {beforeCopyPath}");

                    // -------------------------------------------------
                    // Copy rows: copy row 2 (index 1) to row 4 (index 3)
                    // -------------------------------------------------
                    int sourceRowIndex = 1;        // Row 2 (zero‑based)
                    int destinationRowIndex = 3;   // Row 4 (zero‑based)

                    // Manual row copy because Cells.CopyRows(int, int, int) may not be available in some versions
                    int maxColumn = cells.MaxColumn;
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell sourceCell = cells[sourceRowIndex, col];
                        Cell destCell = cells[destinationRowIndex, col];
                        destCell.Copy(sourceCell);
                    }
                    Console.WriteLine($"Copied row from index {sourceRowIndex} to index {destinationRowIndex}.");

                    // -------------------------------------------------
                    // Recalculate the worksheet so that formulas in the copied rows
                    // (and any other dependent formulas) are updated.
                    // -------------------------------------------------
                    CalculationOptions calcOptions = new CalculationOptions
                    {
                        Recursive = true
                    };

                    worksheet.CalculateFormula(calcOptions, true);
                    Console.WriteLine("Worksheet formulas recalculated.");

                    // Save the final workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Final workbook saved: {outputPath}");
                }

                Console.WriteLine("Process completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
