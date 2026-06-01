using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (including headers)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");

        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Item " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue(i * 10);      // Series1 values
            sheet.Cells[$"C{i}"].PutValue(i * 15);      // Series2 values
        }

        // Add a worksheet table (ListObject) that covers the data range.
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int tableIndex = sheet.ListObjects.Add(0, 0, 6, 3, true);
        var table = sheet.ListObjects[tableIndex];
        table.DisplayName = "DataTable";   // Table name that can be used in formulas

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source to the table columns
        // Category (X‑axis) comes from the "Category" column, series from the other columns
        chart.NSeries.Add("DataTable[Series1]", true);
        chart.NSeries.CategoryData = "DataTable[Category]";

        // Apply a predefined built‑in chart style (valid values: 1‑48)
        chart.Style = 5;   // Example style index

        // Save the workbook
        workbook.Save("ChartWithTableStyle.xlsx", SaveFormat.Xlsx);
    }
}