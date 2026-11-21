// Title: Generate a stacked column chart for quarterly product sales using Aspose.Cells in C#
// Common Searches: how to add a ColumnStacked chart to an Excel workbook with Aspose.Cells C# | Aspose.Cells C# example for creating a stacked column chart with multiple series | C# code to plot cumulative sales by quarter using Aspose.Cells chart API | setting category axis labels for a stacked column chart in Aspose.Cells | saving an Excel file with a stacked column chart using Aspose.Cells in .NET
// Tags: Aspose.Cells create ColumnStacked chart | C# add stacked column series to Excel chart | Aspose.Cells set category axis labels | populate worksheet quarterly sales Aspose.Cells | Aspose.Cells save workbook with chart as XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells A1:D5 with quarter labels and sales figures for three products, adds a ColumnStacked chart, sets the chart title to "Cumulative Sales by Quarter", defines three series for the products, assigns the quarter labels as the category axis, and saves the workbook as StackedColumnChart.xlsx.
class StackedColumnChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: quarters and sales for three products.
            // Row 0: headers
            sheet.Cells["A1"].PutValue("Quarter");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["D1"].PutValue("Product C");

            // Rows 2-5: data for Q1–Q4
            string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
            double[,] sales = {
                { 12000, 15000, 13000 },
                { 14000, 16000, 15000 },
                { 13000, 17000, 16000 },
                { 15000, 18000, 17000 }
            };

            for (int i = 0; i < quarters.Length; i++)
            {
                // Quarter label
                sheet.Cells[i + 1, 0].PutValue(quarters[i]);

                // Sales values for each product
                for (int j = 0; j < 3; j++)
                {
                    sheet.Cells[i + 1, j + 1].PutValue(sales[i, j]);
                }
            }

            // Add a stacked column chart to the worksheet.
            // Note: In Aspose.Cells the enum member is ColumnStacked (not StackedColumn).
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title.
            chart.Title.Text = "Cumulative Sales by Quarter";

            // Add series for each product.
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Product A";

            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Product B";

            chart.NSeries.Add("D2:D5", true);
            chart.NSeries[2].Name = "Product C";

            // Set category (X) axis labels to quarters.
            chart.NSeries.CategoryData = "A2:A5";

            // Save the workbook to a file.
            workbook.Save("StackedColumnChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
