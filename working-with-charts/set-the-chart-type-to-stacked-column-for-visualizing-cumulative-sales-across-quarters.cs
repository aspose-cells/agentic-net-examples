using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

class StackedColumnChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers for quarters and products
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["B1"].PutValue("Product A");
        sheet.Cells["C1"].PutValue("Product B");
        sheet.Cells["D1"].PutValue("Product C");

        // Populate sales data for each quarter
        string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
        int[,] sales = {
            { 120, 150, 100 },
            { 130, 160, 110 },
            { 140, 170, 120 },
            { 150, 180, 130 }
        };

        for (int i = 0; i < quarters.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(quarters[i]);          // Quarter label
            sheet.Cells[i + 1, 1].PutValue(sales[i, 0]);        // Product A
            sheet.Cells[i + 1, 2].PutValue(sales[i, 1]);        // Product B
            sheet.Cells[i + 1, 3].PutValue(sales[i, 2]);        // Product C
        }

        // Add a stacked column chart (ChartType.ColumnStacked)
        int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define series for each product
        chart.NSeries.Add("B2:B5", true);          // Series for Product A
        chart.NSeries[0].Name = "Product A";

        chart.NSeries.Add("C2:C5", true);          // Series for Product B
        chart.NSeries[1].Name = "Product B";

        chart.NSeries.Add("D2:D5", true);          // Series for Product C
        chart.NSeries[2].Name = "Product C";

        // Set category (quarters) for the X‑axis
        chart.NSeries.CategoryData = "A2:A5";

        // Ensure the chart type is stacked column (redundant but explicit)
        chart.Type = ChartType.ColumnStacked;

        // Add a descriptive title
        chart.Title.Text = "Cumulative Sales by Quarter";

        // Save the workbook with the chart
        workbook.Save("StackedColumnChart.xlsx", SaveFormat.Xlsx);
    }
}