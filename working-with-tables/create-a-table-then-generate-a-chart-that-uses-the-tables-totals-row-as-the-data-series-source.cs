// Title: Aspose.Cells C# – Create a Table with a Totals Row and Plot the Total in a Column Chart
// Description: Demonstrates how to build an Excel workbook with a ListObject (A1:B4), enable the totals row, set the Sales column to TotalsCalculation.Sum, and add a column chart whose series is bound directly to the totals cell (e.g., Sheet1!$B$5). The workbook is saved as TableTotalsChart.xlsx.
// Keywords: Aspose.Cells | C# | ListObject | Excel table with totals row | TotalsCalculation.Sum | column chart from totals cell | chart series from cell address | TableTotalsChart example | Aspose.Cells chart binding | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells add totals row to table | C# chart series using totals row cell | Create column chart from summed column value Aspose.Cells | ListObject TotalsCalculation.Sum example | Bind chart to table totals cell Aspose
// Developer Intent: Generate a worksheet table with an automatic sum row and visualize that aggregate value in a column chart using Aspose.Cells for .NET.
// Use Cases: Produce a sales report where the grand total appears as a single column in a chart for instant visual insight. | Automate monthly dashboards by linking table totals directly to charts, removing manual range calculations. | Create concise executive summaries that highlight aggregated metrics from data tables via chart series.
// AI Prompts: Write C# code with Aspose.Cells to create a ListObject, enable a totals row, set TotalsCalculation.Sum for a column, and bind the totals cell to a column chart series. | Explain how to retrieve the address of a totals‑row cell in Aspose.Cells and use it with Chart.NSeries.Add for dynamic chart data. | Suggest formatting options for the totals row and chart title after linking the total value in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsTableTotalsChart
{
    // Demonstrates how to build an Excel workbook with a ListObject (A1:B4), enable the totals row, set the Sales column to TotalsCalculation.Sum, and add a column chart whose series is bound directly to the totals cell (e.g., Sheet1!$B$5). The workbook is saved as TableTotalsChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate header and data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(90);

                // Create a table that includes the header and data rows (A1:B4)
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SalesTable";          // Optional: give the table a friendly name
                table.ShowTotals = true;                    // Enable the totals row

                // Set the totals calculation for the "Sales" column (second column, index 1)
                table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

                // Add a column chart positioned below the table
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Use the totals row cell of the "Sales" column as the data source for the series
                // Totals row is placed immediately after the data range
                int totalsRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount; // zero‑based index
                string totalsCellAddress = sheet.Cells[totalsRowIndex, 1].Name; // column B (index 1)

                // Add the series using the cell address (e.g., =Sheet1!$B$5)
                chart.NSeries.Add($"=Sheet1!{totalsCellAddress}", true);

                // Optional: give the series a name
                chart.NSeries[0].Name = "Total Sales";

                // Save the workbook
                workbook.Save("TableTotalsChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
