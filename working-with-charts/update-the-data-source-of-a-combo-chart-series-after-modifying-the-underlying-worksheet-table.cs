// Title: Refresh a Combo Chart After Expanding a ListObject Table – Aspose.Cells for .NET
// Description: Demonstrates how to bind a column‑line combo chart to a ListObject, add a new row, resize the table, and refresh the chart using Calculate while checking the IsChartDataChanged flag. The workbook is saved as an updated Excel file.
// Keywords: Aspose.Cells | C# combo chart update | ListObject resize | refresh chart after table change | IsChartDataChanged | .NET Excel chart recalculation | structured table chart binding | add row to Excel table Aspose
// Common Searches: how to refresh a combo chart after adding rows in Aspose.Cells | Aspose.Cells update chart data source when ListObject grows | C# recalculate chart after table resize Aspose | check chart data changed flag Aspose.Cells | bind combo chart series to table column .NET
// Developer Intent: Update an existing combo chart so it automatically reflects rows added to the underlying ListObject without recreating the series.
// Use Cases: Automatically extend a sales chart when new records are appended to a structured table. | Detect changes in chart data after table modifications and trigger a recalculation only when needed. | Maintain chart formatting and series types while the source table size changes.
// AI Prompts: Write C# code with Aspose.Cells that adds multiple rows to a ListObject and refreshes all linked chart series. | Explain the purpose of IsChartDataChanged and show how to use it to conditionally recalculate a chart after a table resize. | Provide a step‑by‑step tutorial for binding a combo chart to a ListObject column and keeping it synchronized when the table expands.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsComboChartUpdate
{
    // Demonstrates how to bind a column‑line combo chart to a ListObject, add a new row, resize the table, and refresh the chart using Calculate while checking the IsChartDataChanged flag. The workbook is saved as an updated Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 2. Populate sample data (Category / Value) in cells A1:B5
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B5"].PutValue(40);

                // ------------------------------------------------------------
                // 3. Convert the range into a structured table (ListObject)
                // ------------------------------------------------------------
                int firstRow = 0; // zero‑based index
                int firstColumn = 0;
                int totalRows = 5;   // includes header row
                int totalColumns = 2;
                ListObject table = sheet.ListObjects[sheet.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows - 1, firstColumn + totalColumns - 1, true)];
                table.DisplayName = "SalesTable";

                // ------------------------------------------------------------
                // 4. Add a Combo chart (Column + Line) to the worksheet
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart comboChart = sheet.Charts[chartIndex];
                comboChart.Title.Text = "Sales Combo Chart";

                // ------------------------------------------------------------
                // 5. Add the first series (Column) using the table's data range
                // ------------------------------------------------------------
                string columnValuesRef = "SalesTable[Value]";
                string categoryRef = "SalesTable[Category]";

                comboChart.NSeries.Add($"={sheet.Name}!{columnValuesRef}", true);
                comboChart.NSeries[0].XValues = $"={sheet.Name}!{categoryRef}";
                comboChart.NSeries[0].Type = ChartType.Column;

                // ------------------------------------------------------------
                // 6. Add the second series (Line) using the same data range
                // ------------------------------------------------------------
                comboChart.NSeries.Add($"={sheet.Name}!{columnValuesRef}", true);
                comboChart.NSeries[1].XValues = $"={sheet.Name}!{categoryRef}";
                comboChart.NSeries[1].Type = ChartType.Line;

                // ------------------------------------------------------------
                // 7. Initial calculation of the chart (renders it with the original data)
                // ------------------------------------------------------------
                comboChart.Calculate();

                // ------------------------------------------------------------
                // 8. Modify the underlying table – add a new row (E, 50)
                // ------------------------------------------------------------
                int newRowIndex = sheet.Cells.MaxDataRow + 1; // next empty row
                sheet.Cells[newRowIndex, 0].PutValue("E");    // Category
                sheet.Cells[newRowIndex, 1].PutValue(50);    // Value

                // Resize the table to include the new row.
                // Using the overload that takes row/column indices.
                table.Resize(0, 0, newRowIndex + 1, 2, true);

                // ------------------------------------------------------------
                // 9. Refresh the chart data source after the table change
                // ------------------------------------------------------------
                bool dataChangedBefore = comboChart.IsChartDataChanged(); // expected false
                comboChart.Calculate(); // re‑calculate chart with updated table
                bool dataChangedAfter = comboChart.IsChartDataChanged(); // expected true

                Console.WriteLine($"Chart data changed flag before refresh: {dataChangedBefore}");
                Console.WriteLine($"Chart data changed flag after refresh: {dataChangedAfter}");

                // ------------------------------------------------------------
                // 10. Save the workbook
                // ------------------------------------------------------------
                workbook.Save("ComboChartUpdated.xlsx");
                Console.WriteLine("Workbook saved as ComboChartUpdated.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
