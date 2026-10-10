// Title: Generate a stacked column chart of quarterly product sales with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a workbook, populate quarterly sales data, add a ColumnStacked chart, bind category and series ranges, set a chart title, and save the file as an .xlsx workbook. | Demonstrate how to configure NSeries and CategoryData for a stacked column chart that visualizes cumulative sales of multiple products per quarter.
// Common Searches: how to create a stacked column chart with Aspose.Cells in C# | Aspose.Cells example for cumulative sales chart by quarter | C# code to add ColumnStacked chart to worksheet using Aspose.Cells | binding series data to stacked column chart Aspose.Cells .NET | save workbook with stacked column chart Aspose.Cells
// Tags: Aspose.Cells ColumnStacked chart creation | C# stacked column chart with Aspose.Cells | Aspose.Cells bind series data range | Aspose.Cells generate .xlsx sales chart | Aspose.Cells set chart title programmatically

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Shows how to build a workbook, fill it with quarterly sales figures, add a ColumnStacked chart, assign category labels and product series, set the chart title, and save the result as StackedColumnChart.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data for quarters
            sheet.Cells["A1"].PutValue("Quarter");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["D1"].PutValue("Product C");

            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            sheet.Cells["B2"].PutValue(12000);
            sheet.Cells["B3"].PutValue(15000);
            sheet.Cells["B4"].PutValue(13000);
            sheet.Cells["B5"].PutValue(17000);

            sheet.Cells["C2"].PutValue(10000);
            sheet.Cells["C3"].PutValue(11000);
            sheet.Cells["C4"].PutValue(9000);
            sheet.Cells["C5"].PutValue(12000);

            sheet.Cells["D2"].PutValue(8000);
            sheet.Cells["D3"].PutValue(9500);
            sheet.Cells["D4"].PutValue(10500);
            sheet.Cells["D5"].PutValue(11500);

            // Add a stacked column chart to the worksheet
            // Use ChartType.ColumnStacked (the correct enum value for a stacked column chart)
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Cumulative Sales by Quarter";

            // Define category (X) axis labels (quarters)
            chart.NSeries.CategoryData = "A2:A5";

            // Add series for each product
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Product A";

            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Product B";

            chart.NSeries.Add("D2:D5", true);
            chart.NSeries[2].Name = "Product C";

            // Save the workbook with the chart
            string outputPath = Path.Combine(Environment.CurrentDirectory, "StackedColumnChart.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
