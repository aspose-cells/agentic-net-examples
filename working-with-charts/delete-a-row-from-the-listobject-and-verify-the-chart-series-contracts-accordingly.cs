// Title: Delete a ListObject row and auto‑update a linked column chart with Aspose.Cells (C#)
// Description: Creates a workbook with a table (ListObject) of quarterly data, adds a column chart that references the table, deletes a specific data row using Cells.DeleteRow with reference updates, and demonstrates that the chart series contracts automatically by printing the points count before and after deletion. The workbook is then saved.
// Keywords: Aspose.Cells delete row | ListObject table chart sync | C# chart series update after row removal | Cells.DeleteRow updateReferences | Aspose.Cells column chart auto‑adjust | verify chart points count Aspose.Cells | Aspose.Cells .NET example
// Common Searches: Aspose.Cells delete row from ListObject and update chart | C# remove table row and keep chart data in sync | How to auto‑adjust chart series after deleting rows in Aspose.Cells | Verify chart points count after row deletion Aspose.Cells | Aspose.Cells example for ListObject and chart synchronization
// Developer Intent: Remove a data row from a ListObject and ensure the linked chart reflects the change automatically.
// Use Cases: Delete the Q3 entry from a quarterly sales table and have the column chart display only three points. | Use Cells.DeleteRow with the updateReferences flag so that any chart bound to the table updates its series ranges without manual code. | Save and reopen the workbook to confirm that the chart retains the corrected data range after row removal.
// AI Prompts: Generate C# code with Aspose.Cells that deletes the third data row of a ListObject and automatically updates a linked column chart. | Show how to compare Chart.NSeries[0].Points.Count before and after removing a table row using Aspose.Cells. | Explain the effect of the DeleteRow method's updateReferences parameter on chart data ranges in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook with a table (ListObject) of quarterly data, adds a column chart that references the table, deletes a specific data row using Cells.DeleteRow with reference updates, and demonstrates that the chart series contracts automatically by printing the points count before and after deletion. The workbook is then saved.
    public class DeleteRowFromListObjectAndVerifyChart
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the table (ListObject)
                // Header row
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Series1");
                cells["C1"].PutValue("Series2");

                // Data rows (4 rows of data)
                cells["A2"].PutValue("Q1");
                cells["B2"].PutValue(100);
                cells["C2"].PutValue(150);

                cells["A3"].PutValue("Q2");
                cells["B3"].PutValue(200);
                cells["C3"].PutValue(250);

                cells["A4"].PutValue("Q3");
                cells["B4"].PutValue(300);
                cells["C4"].PutValue(350);

                cells["A5"].PutValue("Q4");
                cells["B5"].PutValue(400);
                cells["C5"].PutValue(450);

                // Create a ListObject (table) that covers the data range including header
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.ShowHeaderRow = true;
                table.ShowTableStyleColumnStripes = true;

                // Add a column chart that uses the table data
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Calculate data range details
                int firstDataRow = table.DataRange.FirstRow + 1;               // zero‑based index of first data row (skip header)
                int dataRowCount = table.DataRange.RowCount - 1;               // number of data rows
                int startRow = firstDataRow + 1;                               // Excel row number (1‑based)
                int endRow = firstDataRow + dataRowCount;                      // Excel row number (1‑based)

                // Series1 (B column)
                string series1Col = CellsHelper.ColumnIndexToName(table.DataRange.FirstColumn + 1);
                string series1Range = $"{series1Col}{startRow}:{series1Col}{endRow}";
                chart.NSeries.Add(series1Range, true);

                // Series2 (C column)
                string series2Col = CellsHelper.ColumnIndexToName(table.DataRange.FirstColumn + 2);
                string series2Range = $"{series2Col}{startRow}:{series2Col}{endRow}";
                chart.NSeries.Add(series2Range, true);

                // Set category (X‑axis) data to the Category column (A column)
                string categoryCol = CellsHelper.ColumnIndexToName(table.DataRange.FirstColumn);
                string categoryRange = $"{categoryCol}{startRow}:{categoryCol}{endRow}";
                chart.NSeries.CategoryData = categoryRange;

                // Display initial points count for the first series
                Console.WriteLine($"Initial points count (Series1): {chart.NSeries[0].Points.Count}");

                // Delete the third data row (Excel row 4, zero‑based index 3)
                try
                {
                    cells.DeleteRow(3, true); // Updates references automatically
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Row deletion error: {ex.Message}");
                }

                // Verify the points count after deletion
                Console.WriteLine($"Points count after row deletion (Series1): {chart.NSeries[0].Points.Count}");

                // Save the workbook
                workbook.Save("DeleteRowFromListObjectAndVerifyChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
