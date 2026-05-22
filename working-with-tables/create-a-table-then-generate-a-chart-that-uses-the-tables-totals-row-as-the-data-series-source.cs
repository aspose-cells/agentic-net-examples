using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

class TableTotalsChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header and sample data (3 data rows)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Create a table (ListObject) that includes the header and data rows (A1:B4)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";          // Optional: give the table a name
            table.ShowTotals = true;                // Enable the totals row

            // Set the totals calculation for the "Value" column (second column, index 1)
            table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

            // Determine the address of the totals row cells for the two columns
            // Totals row is placed immediately after the data rows.
            int totalsRowIndex = table.StartRow + table.DataRange.RowCount; // zero‑based index

            // Column letters for the two columns
            string valueColumnLetter = CellsHelper.ColumnIndexToName(table.StartColumn + 1); // "B"
            string categoryColumnLetter = CellsHelper.ColumnIndexToName(table.StartColumn); // "A"

            // Build the A1‑style references for the totals row cells
            string valueCellRef = $"=Sheet1!${valueColumnLetter}${totalsRowIndex + 1}";
            string categoryCellRef = $"=Sheet1!${categoryColumnLetter}${totalsRowIndex + 1}";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Use the totals row as the data source for the series
            chart.NSeries.Add(valueCellRef, true);          // Series values from totals row
            chart.NSeries[0].Name = "Total";                // Optional series name
            chart.NSeries.CategoryData = categoryCellRef;   // Category label from totals row

            // Optional: set a chart title
            chart.Title.Text = "Totals Row Chart";

            // Save the workbook
            string outputPath = "TableTotalsChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}