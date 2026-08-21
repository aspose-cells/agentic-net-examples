// Title: Delete a ListObject row and auto‑update linked column chart with Aspose.Cells for .NET
// Description: This C# example creates a workbook containing a table (ListObject) and a column chart that uses structured references. It deletes a specific data row from the worksheet, then confirms that the table’s row count and the chart series point counts adjust automatically before saving the file.
// Keywords: Aspose.Cells C# delete ListObject row | update chart series after table row removal | structured reference chart Aspose.Cells | verify column chart points count | Excel table row deletion programmatically
// Common Searches: Aspose.Cells remove row from table and keep chart synced | C# delete ListObject row and refresh chart series | how to verify chart points after deleting a table row in Aspose.Cells
// Developer Intent: Remove a data row from a ListObject and ensure the associated chart reflects the new range.
// Use Cases: Eliminate a quarter’s sales entry from a financial table while the column chart stays accurate. | Clean erroneous rows in a dataset and automatically adjust the visual chart for reporting. | Generate dynamic reports where filtered rows are excluded and the chart must display only remaining data.
// AI Prompts: Write C# code with Aspose.Cells that deletes a given row from a ListObject and updates all linked chart series. | Explain how structured references in Aspose.Cells react when a table row is removed and how to validate the chart’s point count. | Provide step‑by‑step instructions to confirm that a column chart’s NSeries points match the ListObject row count after a deletion.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook containing a table (ListObject) and a column chart that uses structured references. It deletes a specific data row from the worksheet, then confirms that the table’s row count and the chart series point counts adjust automatically before saving the file.
    public class DeleteRowFromListObjectAndVerifyChart
    {
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

                // Data rows (5 rows)
                for (int i = 2; i <= 6; i++)
                {
                    cells[$"A{i}"].PutValue($"Q{i - 1}");
                    cells[$"B{i}"].PutValue(100 * (i - 1));
                    cells[$"C{i}"].PutValue(150 * (i - 1));
                }

                // Create a ListObject (table) that covers the data range A1:C6
                int tableIndex = sheet.ListObjects.Add(0, 0, 5, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.ShowHeaderRow = true;
                table.ShowTableStyleFirstColumn = false;
                table.ShowTableStyleLastColumn = false;
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Add a column chart that uses the table data
                int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Use structured references to set series data ranges
                // Ensure the table has a display name (Aspose.Cells uses DisplayName)
                string tableName = table.DisplayName;

                chart.NSeries.Add($"{tableName}[Series1]", true);
                chart.NSeries.Add($"{tableName}[Series2]", true);
                chart.NSeries.CategoryData = $"{tableName}[Category]";

                // Display initial state
                Console.WriteLine($"Initial row count in table: {table.DataRange.RowCount}");
                Console.WriteLine($"Initial series count: {chart.NSeries.Count}");
                Console.WriteLine($"Initial points in first series: {chart.NSeries[0].Points.Count}");

                // Delete the third data row (index 3 corresponds to Excel row 4)
                cells.DeleteRow(3, true); // Row index 3 = Excel row 4 (Q3)

                // Verify that the table has adjusted
                Console.WriteLine($"\nAfter deletion:");
                Console.WriteLine($"Row count in table: {table.DataRange.RowCount}");

                // Verify that the chart series still reference the correct number of points
                Console.WriteLine($"Series count (should remain unchanged): {chart.NSeries.Count}");
                Console.WriteLine($"Points in first series after deletion: {chart.NSeries[0].Points.Count}");

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeleteRowFromListObjectAndVerifyChart.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the example
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteRowFromListObjectAndVerifyChart.Run();
        }
    }
}
