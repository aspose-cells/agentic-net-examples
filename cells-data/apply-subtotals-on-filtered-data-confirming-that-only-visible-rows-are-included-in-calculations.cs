// Title: Aspose.Cells .NET: Add Subtotal to Filtered Data and Verify Visible Row Sum
// Description: Creates a workbook, fills it with Category/Amount data, applies an AutoFilter to show only rows where Category = "A", inserts a SUM subtotal for the Amount column, manually totals the visible rows, and confirms that the worksheet subtotal matches the calculated visible sum before saving the file.
// Keywords: Aspose.Cells subtotal filtered rows | C# AutoFilter subtotal | visible rows sum Aspose.Cells | Cells.Subtotal method .NET | retrieve subtotal setting | Excel subtotal filtered data | calculate visible sum C# | filter and subtotal Aspose.Cells
// Common Searches: Aspose.Cells add subtotal to filtered range | sum only visible rows after AutoFilter Aspose.Cells | retrieve subtotal row value using Aspose.Cells | validate subtotal matches visible rows .NET | C# code for subtotal with AutoFilter
// Developer Intent: Insert a subtotal for a column in a filtered worksheet and ensure the calculation includes only the rows that remain visible.
// Use Cases: Generate a financial report that displays selected categories and automatically adds a subtotal for the visible amounts. | Export data to Excel, apply an AutoFilter, compute grouped subtotals, and programmatically verify the results against a manual sum of visible rows. | Extract the subtotal row and its value for further processing in a .NET application, such as displaying a summary in a UI.
// AI Prompts: Show how to apply a subtotal to a filtered range in Aspose.Cells and retrieve the inserted subtotal row. | Provide C# code that sums only the visible rows after an AutoFilter and compares the result with the worksheet subtotal. | Explain the Subtotal parameters (replace existing, page breaks, summary below) in Aspose.Cells and how they affect filtered data.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSubtotalFilteredDemo
{
    // Creates a workbook, fills it with Category/Amount data, applies an AutoFilter to show only rows where Category = "A", inserts a SUM subtotal for the Amount column, manually totals the visible rows, and confirms that the worksheet subtotal matches the calculated visible sum before saving the file.
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

                // Populate sample data
                // Header
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");

                // Data rows
                object[,] data = new object[,]
                {
                    { "A", 100 },
                    { "B", 200 },
                    { "A", 150 },
                    { "B", 250 },
                    { "A", 300 },
                    { "B", 350 }
                };

                for (int r = 0; r < data.GetLength(0); r++)
                {
                    cells[r + 1, 0].PutValue(data[r, 0]); // Category
                    cells[r + 1, 1].PutValue(data[r, 1]); // Amount
                }

                // Apply AutoFilter on the header row (A1:B7)
                sheet.AutoFilter.Range = "A1:B7";

                // Filter to show only rows where Category = "A"
                sheet.AutoFilter.AddFilter(0, "A");
                // Refresh filter (hide rows that do not meet the criteria)
                sheet.AutoFilter.Refresh();

                // Define the cell area that includes the header and all original data rows
                // EndRow is the last row of original data (row index 6, zero‑based)
                CellArea area = CellArea.CreateCellArea(0, 0, 6, 1);

                // Apply subtotal: group by Category (column 0), sum Amount (column 1)
                // Parameters: replace existing subtotals = false, page breaks = false, summary below data = false
                cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, false, false, false);

                // Retrieve subtotal setting (optional, just to demonstrate the rule)
                SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
                Console.WriteLine($"Subtotal Function: {setting.SubtotalFunction}");
                Console.WriteLine($"GroupBy column index: {setting.GroupBy}");
                Console.WriteLine($"SummaryBelowData: {setting.SummaryBelowData}");

                // Calculate the sum of visible rows manually
                double visibleSum = 0;
                // Original data rows are from row 1 to row 6 (zero‑based)
                for (int row = 1; row <= 6; row++)
                {
                    if (!cells.IsRowHidden(row))
                    {
                        visibleSum += cells[row, 1].DoubleValue;
                    }
                }
                Console.WriteLine($"Manual sum of visible rows: {visibleSum}");

                // Locate the subtotal row inserted by the Subtotal method
                // The subtotal row contains the label "Sum" in the first column
                int subtotalRow = -1;
                int maxRow = cells.MaxDataRow;
                for (int row = 0; row <= maxRow; row++)
                {
                    if (cells[row, 0].StringValue.Equals("Sum", StringComparison.OrdinalIgnoreCase))
                    {
                        subtotalRow = row;
                        break;
                    }
                }

                if (subtotalRow != -1)
                {
                    double subtotalValue = cells[subtotalRow, 1].DoubleValue;
                    Console.WriteLine($"Subtotal value from worksheet: {subtotalValue}");
                    Console.WriteLine($"Subtotal matches manual visible sum: {Math.Abs(subtotalValue - visibleSum) < 0.0001}");
                }
                else
                {
                    Console.WriteLine("Subtotal row not found.");
                }

                // Save the workbook (ensure the directory exists)
                string outputPath = "SubtotalFilteredDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
