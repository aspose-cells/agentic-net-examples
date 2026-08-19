// Title: C# – Update Combo Chart Series After Expanding an Excel Table Using Aspose.Cells
// Description: Demonstrates how to create a ListObject (Excel table), add a column‑line combo chart that uses a structured reference for the sales series and a range reference for a target series, append new rows, resize the table, rebuild the chart series to include the new data, force chart recalculation, and save both the original and updated workbooks.
// Keywords: Aspose.Cells | .NET | C# | combo chart | update chart series | Excel table resize | ListObject | structured reference | chart recalculation | dynamic data source
// Common Searches: Aspose.Cells update chart after adding rows to ListObject | C# refresh combo chart data range when Excel table grows | Resize Excel table and keep chart linked Aspose.Cells | Change chart series source to structured reference .NET | Recalculate chart after modifying worksheet data Aspose
// Developer Intent: Refresh the combo chart so its series automatically include rows added to the underlying Excel table.
// Use Cases: Add new monthly records to a ListObject and keep the column series linked via a structured reference. | Extend a hidden target column, update the line series range, and recalculate the chart to display new target values. | Resize an Excel table after data insertion, clear existing NSeries, re‑add them with updated ranges, and save the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that appends rows to a ListObject and updates a combo chart’s series ranges. | Explain how structured references keep a chart series linked after resizing an Excel table in Aspose.Cells. | Show the steps to clear and rebuild chart series in Aspose.Cells after modifying worksheet data.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsComboChartUpdate
{
    // Demonstrates how to create a ListObject (Excel table), add a column‑line combo chart that uses a structured reference for the sales series and a range reference for a target series, append new rows, resize the table, rebuild the chart series to include the new data, force chart recalculation, and save both the original and updated workbooks.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate initial data (A1:B6) – this will be the source table
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                int[] sales = { 120, 150, 180, 200, 170 };
                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A (zero‑based row index)
                    sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                }

                // Convert the range into a ListObject (Excel Table) for easier reference
                int tableIndex = sheet.ListObjects.Add(0, 0, 6, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Set the table name (use DisplayName for compatibility)
                table.DisplayName = "SalesTable";

                // Add a Combo chart (Column + Line) to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Monthly Sales (Combo)";

                // First series – Column (Sales) using structured reference
                chart.NSeries.Add("=Sheet1!SalesTable[Sales]", true);
                chart.NSeries[0].Name = "Sales";

                // Second series – Line (Target values placed in hidden column C)
                for (int i = 0; i < sales.Length; i++)
                {
                    sheet.Cells[i + 1, 2].PutValue(sales[i] * 1.1);
                }
                chart.NSeries.Add("=Sheet1!C2:C6", true);
                chart.NSeries[1].Name = "Target";
                chart.NSeries[1].Type = ChartType.Line;

                // Save the initial workbook
                string initialPath = "ComboChart_Initial.xlsx";
                workbook.Save(initialPath);
                Console.WriteLine($"Saved: {Path.GetFullPath(initialPath)}");

                // ------------------------------------------------------------
                // MODIFY THE UNDERLYING TABLE: add new month data
                // ------------------------------------------------------------
                // Determine the first empty row after the existing table data
                int firstNewRow = sheet.Cells.MaxDataRow + 1; // zero‑based index

                // Add June
                sheet.Cells[firstNewRow, 0].PutValue("Jun");
                sheet.Cells[firstNewRow, 1].PutValue(190);
                sheet.Cells[firstNewRow, 2].PutValue(190 * 1.1);

                // Add July
                sheet.Cells[firstNewRow + 1, 0].PutValue("Jul");
                sheet.Cells[firstNewRow + 1, 1].PutValue(210);
                sheet.Cells[firstNewRow + 1, 2].PutValue(210 * 1.1);

                // Resize the table to include the new rows (header + 7 data rows = 8 total rows)
                // hasHeaders = true because the first row contains column names
                table.Resize(0, 0, 8, 2, true);

                // ------------------------------------------------------------
                // UPDATE CHART SERIES DATA RANGES to reflect the expanded table
                // ------------------------------------------------------------
                chart.NSeries.Clear();

                // Sales series (column) – still uses structured reference
                chart.NSeries.Add("=Sheet1!SalesTable[Sales]", true);
                chart.NSeries[0].Name = "Sales";

                // Target series (line) – updated range includes new rows
                int lastTargetRowNumber = firstNewRow + 2; // Excel row number (1‑based)
                chart.NSeries.Add($"=Sheet1!C2:C{lastTargetRowNumber}", true);
                chart.NSeries[1].Name = "Target";
                chart.NSeries[1].Type = ChartType.Line;

                // Force chart recalculation
                chart.Calculate();

                // Save the workbook after updating the chart
                string updatedPath = "ComboChart_Updated.xlsx";
                workbook.Save(updatedPath);
                Console.WriteLine($"Saved: {Path.GetFullPath(updatedPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
