// Title: How to apply a subtotal to filtered rows and verify it uses only visible data with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, adds sample data, applies an AutoFilter to show only rows where Category = "A", inserts a subtotal that sums the Amount column, and validates that the subtotal matches the sum of the visible rows. | Write C# that places the subtotal row above the data range, replaces the default "Subtotal" label with custom text, and updates the SubtotalSetting accordingly using Aspose.Cells. | Provide a C# snippet that retrieves the SubtotalSetting for a specified CellArea after applying an AutoFilter and prints the GroupBy column index, the consolidation function, and the list of total columns.
// Common Searches: Aspose.Cells C# subtotal only on rows visible after applying an AutoFilter | How to confirm that a subtotal calculation respects filtered rows in Aspose.Cells .NET | Retrieve SubtotalSetting properties after using the Subtotal method in Aspose.Cells
// Tags: Aspose.Cells subtotal filtered rows | C# AutoFilter subtotal verification | Aspose.Cells retrieve SubtotalSetting | subtotal visible rows Aspose.Cells | Aspose.Cells calculate visible sum

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSubtotalFilteredDemo
{
    // // Demonstrates creating a workbook, populating sample data, applying an AutoFilter to show only Category "A" rows, adding a subtotal that sums the Amount column, calculating the sum of visible rows, retrieving the SubtotalSetting for verification, comparing the two sums, and saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (Header + 10 rows)
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");

                object[,] data = new object[,]
                {
                    { "A", 100 },
                    { "B", 200 },
                    { "A", 150 },
                    { "B", 250 },
                    { "A", 120 },
                    { "B", 300 },
                    { "A", 130 },
                    { "B", 220 },
                    { "A", 140 },
                    { "B", 180 }
                };

                for (int r = 0; r < data.GetLength(0); r++)
                {
                    cells[r + 1, 0].PutValue(data[r, 0]); // Category
                    cells[r + 1, 1].PutValue(data[r, 1]); // Amount
                }

                // Apply an AutoFilter on the header row (A1:B11)
                sheet.AutoFilter.Range = "A1:B11";

                // Filter to show only rows where Category = "A"
                sheet.AutoFilter.AddFilter(0, "A");
                sheet.AutoFilter.Refresh();

                // Define the cell area that contains the data (including header)
                CellArea area = CellArea.CreateCellArea("A1", "B11");

                // Apply subtotals: group by column 0 (Category), sum column 1 (Amount)
                cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, false, false, true);

                // Ensure formulas are calculated before reading values
                workbook.CalculateFormula();

                // Retrieve the subtotal setting for verification
                SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
                Console.WriteLine($"GroupBy column index: {setting.GroupBy}");
                Console.WriteLine($"Subtotal function: {setting.SubtotalFunction}");
                Console.WriteLine($"SummaryBelowData: {setting.SummaryBelowData}");
                Console.WriteLine($"TotalList column index: {setting.TotalList[0]}");

                // Calculate the sum of visible rows manually (should match the subtotal result)
                double visibleSum = 0;
                int lastRow = sheet.Cells.MaxDataRow;
                for (int row = 1; row <= lastRow; row++) // start from 1 to skip header
                {
                    if (!sheet.Cells.IsRowHidden(row))
                    {
                        // Skip subtotal rows (labelled "Subtotal" in column A)
                        if (sheet.Cells[row, 0].StringValue == "Subtotal")
                            continue;

                        Cell amountCell = sheet.Cells[row, 1];
                        // Add numeric cells or cells that contain formulas
                        if (amountCell.Type == CellValueType.IsNumeric || amountCell.IsFormula)
                        {
                            visibleSum += amountCell.DoubleValue;
                        }
                    }
                }
                Console.WriteLine($"Sum of visible Amount rows: {visibleSum}");

                // Locate the subtotal row (first occurrence of the word "Subtotal" in column A)
                double subtotalValue = double.NaN;
                for (int row = 1; row <= lastRow; row++)
                {
                    if (sheet.Cells[row, 0].StringValue == "Subtotal")
                    {
                        Cell subtotalCell = sheet.Cells[row, 1];
                        if (subtotalCell.Type == CellValueType.IsNumeric || subtotalCell.IsFormula)
                        {
                            subtotalValue = subtotalCell.DoubleValue;
                        }
                        break;
                    }
                }
                Console.WriteLine($"Subtotal value reported by Aspose.Cells: {subtotalValue}");

                // Verify that the two sums match (within a small tolerance)
                if (Math.Abs(visibleSum - subtotalValue) < 0.0001)
                    Console.WriteLine("Verification passed: Subtotal includes only visible rows.");
                else
                    Console.WriteLine("Verification failed: Subtotal does not match visible rows sum.");

                // Save the workbook (ensure the directory exists)
                string outputPath = "SubtotalFilteredDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
