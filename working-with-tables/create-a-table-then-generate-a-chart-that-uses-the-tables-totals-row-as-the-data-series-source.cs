// Title: Generate an Excel table with a totals row and create a column chart from the totals using Aspose.Cells for .NET (C#)
// Common Searches: how to add a totals row to an Excel table and use it as a chart source with Aspose.Cells in C# | Aspose.Cells C# create ListObject, enable totals row, and plot column chart from totals | C# code to sum column values in an Excel table and chart the sums using Aspose.Cells | using Aspose.Cells to generate a column chart from the totals row of a worksheet table | Aspose.Cells .NET example for adding SUM formulas to a table and charting the totals
// Tags: Aspose.Cells create ListObject with totals row | Aspose.Cells add SUM formulas to table columns | Aspose.Cells generate column chart from totals row | Aspose.Cells C# chart data series from table totals | Aspose.Cells export workbook to XLSX with chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;   // Required for ListObject (table) support

// The example creates a new workbook, fills it with product and quarterly data, converts the range into a ListObject (Excel table) with a totals row that uses SUM formulas, then adds a column chart whose series values come from the totals row and whose categories are the header labels. The workbook is saved as TableWithTotalsChart.xlsx.
class TableChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill sample data (including header)
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Q1");
            cells["C1"].PutValue("Q2");

            cells["A2"].PutValue("A");
            cells["B2"].PutValue(120);
            cells["C2"].PutValue(150);

            cells["A3"].PutValue("B");
            cells["B3"].PutValue(200);
            cells["C3"].PutValue(180);

            cells["A4"].PutValue("C");
            cells["B4"].PutValue(160);
            cells["C4"].PutValue(170);

            cells["A5"].PutValue("D");
            cells["B5"].PutValue(140);
            cells["C5"].PutValue(190);

            // Define the range that will become a table (including header)
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // column A
            int totalRows = 5;  // header + 4 data rows
            int totalCols = 3;  // columns A‑C

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol,
                firstRow + totalRows - 1, firstCol + totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.ShowTotals = true; // Enable totals row

            // Calculate the index of the totals row (zero‑based)
            int totalsRowIndex = firstRow + totalRows;

            // Set totals formulas for numeric columns (Q1 and Q2)
            // Column B (Q1) total
            cells[totalsRowIndex, 1].Formula = $"=SUM({cells[firstRow + 1, 1].Name}:{cells[firstRow + totalRows - 1, 1].Name})";
            // Column C (Q2) total
            cells[totalsRowIndex, 2].Formula = $"=SUM({cells[firstRow + 1, 2].Name}:{cells[firstRow + totalRows - 1, 2].Name})";

            // Add a column chart that uses the totals row as the data series source
            int chartUpperRow = 7;   // zero‑based index for Excel row 8
            int chartUpperCol = 0;   // column A
            int chartLowerRow = 19;  // Excel row 20
            int chartLowerCol = 6;   // column G

            int chartIdx = sheet.Charts.Add(ChartType.Column, chartUpperRow, chartUpperCol,
                chartLowerRow, chartLowerCol);
            Chart chart = sheet.Charts[chartIdx];
            chart.Title.Text = "Totals per Quarter";

            // Define the series values range (totals row, columns B and C)
            // Totals row is the row immediately after the data rows (Excel row 6)
            string totalsRange = "B6:C6";
            chart.NSeries.Add(totalsRange, false); // false = use categories from separate range

            // Use the header row as category labels (Q1 and Q2)
            chart.NSeries.CategoryData = "B1:C1";

            // Optional: set the series name
            chart.NSeries[0].Name = "Totals";

            // Save the workbook to a file
            string outputPath = "TableWithTotalsChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
