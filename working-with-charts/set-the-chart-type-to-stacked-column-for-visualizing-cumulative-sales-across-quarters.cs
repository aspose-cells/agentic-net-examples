using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace StackedColumnChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data for four quarters
            // Row 0: headers
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Q1");
            sheet.Cells["C1"].PutValue("Q2");
            sheet.Cells["D1"].PutValue("Q3");
            sheet.Cells["E1"].PutValue("Q4");

            // Row 1: Product A
            sheet.Cells["A2"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(150);
            sheet.Cells["D2"].PutValue(130);
            sheet.Cells["E2"].PutValue(170);

            // Row 2: Product B
            sheet.Cells["A3"].PutValue("Product B");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["C3"].PutValue(90);
            sheet.Cells["D3"].PutValue(100);
            sheet.Cells["E3"].PutValue(110);

            // Row 3: Product C
            sheet.Cells["A4"].PutValue("Product C");
            sheet.Cells["B4"].PutValue(60);
            sheet.Cells["C4"].PutValue(70);
            sheet.Cells["D4"].PutValue(65);
            sheet.Cells["E4"].PutValue(75);

            // Add a stacked column chart to visualize cumulative sales
            // Parameters: chart type, top row, left column, bottom row, right column
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the series (sales values) and categories (products)
            chart.NSeries.Add("B2:E4", true);          // Series data (by column)
            chart.NSeries.CategoryData = "A2:A4";      // Category labels (product names)

            // Set a descriptive title
            chart.Title.Text = "Cumulative Quarterly Sales";

            // Save the workbook with the chart
            workbook.Save("StackedColumnSales.xlsx", SaveFormat.Xlsx);
        }
    }
}