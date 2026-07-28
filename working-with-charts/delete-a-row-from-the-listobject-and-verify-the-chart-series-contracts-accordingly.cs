// Title: Delete a ListObject row and automatically update linked chart series with Aspose.Cells (C#)
// Description: This C# example creates a workbook, adds a ListObject (table) and a column chart bound to the table columns, removes a specific data row with Cells.DeleteRow, and then confirms that the table reports the reduced row count and the chart series reflect the new data range before saving the file.
// Keywords: Aspose.Cells delete row | ListObject table row removal | chart series update after row deletion | C# Aspose.Cells chart binding | verify chart data range | Cells.DeleteRow ListObject | Aspose.Cells NSeries point count
// Common Searches: Aspose.Cells delete row from ListObject and keep chart in sync | C# verify chart series after removing a table row | How to refresh chart data range after deleting a row in Aspose.Cells | Remove data row from Excel table and update chart using Aspose.Cells | Aspose.Cells example for ListObject row deletion and chart validation
// Developer Intent: Remove a data row from a ListObject and ensure the associated chart automatically reflects the new range.
// Use Cases: Programmatically delete a row from an Excel table while preserving chart bindings. | Validate that ListObject.DataRange.RowCount decreases after row removal. | Check that Chart.NSeries.Count remains unchanged but Points.Count updates to match the trimmed table. | Display the updated series values directly from worksheet cells. | Save the modified workbook for further processing or download.
// AI Prompts: Write C# code using Aspose.Cells to delete the third data row from a ListObject and print the updated point count for each chart series. | Show how to confirm that a column chart bound to a table automatically adjusts its data range after a row is removed. | Explain step‑by‑step how Cells.DeleteRow impacts ListObject.DataRange and Chart.NSeries in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds a ListObject (table) and a column chart bound to the table columns, removes a specific data row with Cells.DeleteRow, and then confirms that the table reports the reduced row count and the chart series reflect the new data range before saving the file.
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

                // Populate sample data with a header row
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Series1");
                cells["C1"].PutValue("Series2");

                // Data rows (5 rows)
                for (int i = 0; i < 5; i++)
                {
                    cells[$"A{i + 2}"].PutValue($"Item{i + 1}");
                    cells[$"B{i + 2}"].PutValue(10 * (i + 1));   // Series1 values: 10,20,...
                    cells[$"C{i + 2}"].PutValue(15 * (i + 1));   // Series2 values: 15,30,...
                }

                // Create a ListObject (table) that covers the data including header
                int tableIndex = sheet.ListObjects.Add(0, 0, 5, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "DataTable";

                // Add a column chart that uses the table data
                int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Use the table's column ranges for the series
                chart.NSeries.Add($"{table.DisplayName}[Series1]", true);
                chart.NSeries.Add($"{table.DisplayName}[Series2]", true);
                chart.NSeries.CategoryData = $"{table.DisplayName}[Category]";

                Console.WriteLine($"Initial chart series count: {chart.NSeries.Count}");
                Console.WriteLine($"Initial points in first series: {chart.NSeries[0].Points.Count}");

                // Delete the third data row (zero‑based index 3, Excel row 4)
                cells.DeleteRow(3);

                // Verify that the ListObject has one less data row
                Console.WriteLine($"Table data rows after deletion: {table.DataRange.RowCount - 1}");

                // Verify that the chart series still exists and reflects the updated data range
                Console.WriteLine($"Chart series count after row deletion: {chart.NSeries.Count}");
                Console.WriteLine($"Points in first series after deletion: {chart.NSeries[0].Points.Count}");

                // Display the new values of the first series by reading the worksheet cells directly
                Console.WriteLine("Values of first series after deletion:");
                int firstDataRow = table.DataRange.FirstRow + 1; // skip header
                int lastDataRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1;
                int series1Column = table.DataRange.FirstColumn + 1; // column B (Series1)

                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    Console.WriteLine(cells[row, series1Column].Value);
                }

                // Save the workbook
                string outputPath = "DeleteRowFromListObjectAndVerifyChart.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteRowFromListObjectAndVerifyChart.Run();
        }
    }
}
