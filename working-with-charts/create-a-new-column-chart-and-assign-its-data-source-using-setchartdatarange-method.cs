// Title: How to create a column chart and bind its data range with SetChartDataRange using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart data range for column chart example | Create column chart in Excel workbook using Aspose.Cells and SetChartDataRange | Bind worksheet range to a column chart with Aspose.Cells .NET API | Set chart title and data source for column chart in Aspose.Cells C# | Save Excel file containing a column chart using Aspose.Cells for .NET
// Tags: Aspose.Cells SetChartDataRange column chart | C# create Excel column chart Aspose.Cells | binding worksheet range to chart Aspose.Cells | setting chart title Aspose.Cells .NET | saving workbook with chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample code demonstrates how to generate an Excel workbook with Aspose.Cells, populate cells A1:B5 with category labels and numeric values, add a column chart, bind the data range A2:B5 to the chart using SetChartDataRange, set a chart title, and save the workbook as ColumnChart.xlsx.
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

            // Populate sample data for the column chart
            // Column A: Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column B: Values
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(25);

            // Add a new column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (categories and values)
            chart.SetChartDataRange("A2:B5", true);

            // Set chart title
            chart.Title.Text = "Monthly Sales";

            // Save the workbook
            string outputPath = "ColumnChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
