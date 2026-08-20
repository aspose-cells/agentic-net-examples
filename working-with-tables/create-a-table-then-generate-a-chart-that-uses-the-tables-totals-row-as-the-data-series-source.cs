// Title: Aspose.Cells .NET – Create a Table with Totals Row and Plot the Total in a Column Chart
// Description: C# example that builds a new workbook, inserts a two‑column ListObject (Product / Sales), enables the table's Totals row with a SUM calculation, and adds a column chart whose series points to the absolute totals cell (e.g., B5). The workbook is saved as TableWithTotalsChart.xlsx.
// Keywords: Aspose.Cells table totals row | ListObject totals calculation sum | chart from table totals row C# | Aspose.Cells column chart data source | create Excel table with totals Aspose.Cells | .NET Excel chart from totals row
// Common Searches: Aspose.Cells add totals row to ListObject | use table totals row as chart source Aspose.Cells | C# create column chart from table total | Aspose.Cells set totals calculation sum | reference totals row in chart series formula
// Developer Intent: Generate an Excel file that contains a ListObject with a summed totals row and a column chart that visualizes that total.
// Use Cases: Summarize sales figures and display the grand total as a single column in a report. | Automate financial dashboards where table totals drive chart visuals. | Create recurring Excel reports that pull calculated totals directly into charts without manual formulas.
// AI Prompts: Write C# code using Aspose.Cells to add a ListObject with a totals row that calculates the average of a column and generate a pie chart referencing the average cell. | Show how to change the totals calculation to Count and update the chart series to reflect the new totals cell. | Explain how to programmatically locate the totals row index when the table size changes and bind it to a chart series in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsTableChartTotals
{
    // C# example that builds a new workbook, inserts a two‑column ListObject (Product / Sales), enables the table's Totals row with a SUM calculation, and adds a column chart whose series points to the absolute totals cell (e.g., B5). The workbook is saved as TableWithTotalsChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate data ----------
            // Header
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");

            // Data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(90);

            // ---------- Create a table (ListObject) ----------
            // The range includes header (row 0) and data rows (rows 1‑3)
            int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";

            // Enable the totals row and set the calculation for the Sales column (index 1)
            table.ShowTotals = true;
            table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

            // After enabling totals, the totals row is placed just below the data rows.
            // In this example it will be at Excel row 5 (cell B5).
            // ---------- Create a chart that uses the totals row as its data source ----------
            // Add a column chart somewhere below the table
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series source to the totals cell (B5). The formula must be an absolute reference.
            chart.NSeries.Add("=Sheet1!$B$5", true);

            // Optionally set a title for clarity
            chart.Title.Text = "Total Sales (from Table Totals Row)";

            // ---------- Save the workbook ----------
            workbook.Save("TableWithTotalsChart.xlsx");
        }
    }
}
