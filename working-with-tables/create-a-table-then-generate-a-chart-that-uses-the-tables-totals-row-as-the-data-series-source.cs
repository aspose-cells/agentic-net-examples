// Title: Generate an Excel table with a totals row and create a column chart from the totals row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add totals row to ListObject and use it as chart data source | Create column chart from table totals row with Aspose.Cells .NET | How to reference totals row cells in a chart using Aspose.Cells C# | Generate Excel chart from summed columns in a table using Aspose.Cells | C# Aspose.Cells example for totals row and column chart
// Tags: Aspose.Cells ListObject totals row | Aspose.Cells column chart from totals | C# create Excel table with totals | Aspose.Cells chart series range | Aspose.Cells generate chart from table

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

// The sample program creates a new workbook, adds a ListObject (Excel table) with a totals row that sums two value columns, enables the totals row, and then builds a column chart whose series reference the summed cells in the totals row. Category labels are taken from the first column, and the workbook is saved as TableWithTotalsChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1:C5)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value1");
            sheet.Cells["C1"].PutValue("Value2");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Define the range for the table (including header row)
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // zero‑based index for column A
            int totalRows = 5;  // rows 1‑5 (including header)
            int totalCols = 3;  // columns A‑C

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol,
                firstRow + totalRows - 1, firstCol + totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Enable the totals row and set sum calculations for columns B and C
            table.ShowTotals = true;
            table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum; // Value1
            table.ListColumns[2].TotalsCalculation = TotalsCalculation.Sum; // Value2

            // Row index of the totals row (zero‑based)
            int totalsRowIndex = firstRow + totalRows;

            // Create a column chart anchored to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Totals Row Chart";

            // Series 1: totals of Value1 (column B)
            string series1Range = $"'{sheet.Name}'!B{totalsRowIndex + 1}";
            chart.NSeries.Add(series1Range, true);
            chart.NSeries[0].Name = "Total Value1";

            // Series 2: totals of Value2 (column C)
            string series2Range = $"'{sheet.Name}'!C{totalsRowIndex + 1}";
            chart.NSeries.Add(series2Range, true);
            chart.NSeries[1].Name = "Total Value2";

            // Category axis labels (use the header of the first column)
            string categoryRange = $"'{sheet.Name}'!A{firstRow + 1}:A{firstRow + totalRows}";
            chart.NSeries.CategoryData = categoryRange;

            // Save the workbook
            workbook.Save("TableWithTotalsChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
