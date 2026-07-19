// Title: C# – Update Combo Chart Series After Expanding a ListObject Table with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a structured table, adds a column‑line combo chart that references the table, appends a new row, resizes the ListObject, resets the chart series ranges, recalculates the chart, and saves the workbook. It demonstrates how to keep a combo chart synchronized with dynamic table data.
// Keywords: Aspose.Cells C# combo chart | update chart series range | ListObject resize | expand worksheet table | refresh chart after adding rows | structured table chart integration | Aspose.Cells chart Calculate | dynamic data source Aspose.Cells
// Common Searches: Aspose.Cells update combo chart after adding rows | Resize ListObject and refresh chart C# | Change chart series source programmatically Aspose.Cells | Add new row to table and update chart Aspose.Cells | C# example for dynamic combo chart with ListObject
// Developer Intent: Synchronize a combo chart with a ListObject after the table has been expanded.
// Use Cases: Automatically extend a sales table each quarter and have the combo chart display the new data without manual range edits. | Build a live dashboard where chart series adapt to rows added by an automated data import process. | Create a reporting utility that adds periods to a structured table and instantly refreshes the associated column‑line chart.
// AI Prompts: Write C# code using Aspose.Cells to add a row to an existing ListObject, resize the table, and update all combo chart series ranges. | Show how to programmatically change the data source of a combo chart's column and line series after expanding the source table in Aspose.Cells for .NET. | Provide a step‑by‑step Aspose.Cells example that resizes a ListObject and calls Chart.Calculate() to refresh a combo chart with new rows.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsComboChartUpdate
{
    // This Aspose.Cells for .NET example creates a structured table, adds a column‑line combo chart that references the table, appends a new row, resizes the ListObject, resets the chart series ranges, recalculates the chart, and saves the workbook. It demonstrates how to keep a combo chart synchronized with dynamic table data.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate worksheet with sample data ----------
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Sales 2022");
                sheet.Cells["C1"].PutValue("Sales 2023");

                // Data rows
                string[] categories = { "Q1", "Q2", "Q3", "Q4" };
                int[] sales2022 = { 120, 150, 130, 170 };
                int[] sales2023 = { 140, 160, 150, 180 };

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2; // Data starts at row 2 (1‑based)
                    sheet.Cells[row, 0].PutValue(categories[i]);   // Column A
                    sheet.Cells[row, 1].PutValue(sales2022[i]);    // Column B
                    sheet.Cells[row, 2].PutValue(sales2023[i]);    // Column C
                }

                // ---------- Convert the range into a structured table ----------
                int firstRow = 0;          // zero‑based index
                int firstColumn = 0;
                int totalRows = categories.Length + 1; // +1 for header
                int totalColumns = 3;
                int listObjectIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
                ListObject dataTable = sheet.ListObjects[listObjectIndex];
                dataTable.DisplayName = "SalesData";

                // ---------- Add a Combo chart (Column + Line) ----------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 26, 15);
                Chart comboChart = sheet.Charts[chartIndex];
                comboChart.Title.Text = "Sales Comparison";

                // First series – Column (2022)
                comboChart.NSeries.Add("=SalesData!$B$2:$B$5", true);
                comboChart.NSeries[0].Name = "2022 Sales";

                // Second series – Line (2023)
                comboChart.NSeries.Add("=SalesData!$C$2:$C$5", true);
                comboChart.NSeries[1].Name = "2023 Sales";
                comboChart.NSeries[1].Type = ChartType.Line; // Change series type to Line

                // ---------- Modify the underlying table ----------
                // Add a new quarter (Q5) with sales data
                int newRowIndex = sheet.Cells.MaxDataRow + 1; // Append after existing data
                sheet.Cells[newRowIndex, 0].PutValue("Q5");
                sheet.Cells[newRowIndex, 1].PutValue(190); // 2022
                sheet.Cells[newRowIndex, 2].PutValue(200); // 2023

                // Expand the ListObject to include the new row
                int tblFirstRow = dataTable.DataRange.FirstRow;
                int tblFirstColumn = dataTable.DataRange.FirstColumn;
                int tblRowCount = dataTable.DataRange.RowCount + 1; // add one row
                int tblColumnCount = dataTable.DataRange.ColumnCount;
                // The Resize method requires a hasHeaders flag; the table has headers.
                dataTable.Resize(tblFirstRow, tblFirstColumn, tblRowCount, tblColumnCount, true);

                // ---------- Update the chart series data source ----------
                // Explicitly set new ranges to include the added row
                comboChart.NSeries[0].Values = "=SalesData!$B$2:$B$6";
                comboChart.NSeries[1].Values = "=SalesData!$C$2:$C$6";

                // Re‑calculate the chart so that it picks up the new data.
                comboChart.Calculate();

                // ---------- Save the workbook ----------
                string outputPath = "ComboChartUpdated.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
