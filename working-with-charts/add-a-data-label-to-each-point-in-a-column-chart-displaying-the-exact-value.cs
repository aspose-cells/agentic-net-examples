// Title: C# – Add Exact Value Data Labels to a Column Chart Using Aspose.Cells
// Description: Demonstrates how to create a workbook, populate category and numeric data, insert a column chart, bind the series, enable data labels, set each label to the point's exact value, and save the file as an .xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells column chart data labels C# | add value labels to chart points Aspose | C# Aspose.Cells show exact values on columns | Excel chart data labels Aspose.Cells | Aspose.Cells series DataLabels ShowValue | custom data label text Aspose.Cells | .NET chart labeling example
// Common Searches: Aspose.Cells display values on each column in a chart C# | How to enable data labels for a column chart with Aspose.Cells | Set custom data label text for chart points Aspose.Cells | C# add exact value labels to Excel column chart using Aspose | Aspose.Cells column chart label formatting
// Developer Intent: Add data labels that show the precise numeric value for every column in an Aspose.Cells chart.
// Use Cases: Sales dashboards where each bar displays its revenue figure directly on the chart. | Performance reports that label KPI columns with exact metrics for instant readability. | Financial statements exported to Excel with column charts that include value labels to avoid mouse‑over lookups.
// AI Prompts: Generate C# code with Aspose.Cells that adds value data labels to each point of a column chart and formats the label font. | Explain how to position data labels above, inside, or outside columns in an Aspose.Cells chart. | Provide a snippet that hides data labels for columns whose values are below a given threshold using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Demonstrates how to create a workbook, populate category and numeric data, insert a column chart, bind the series, enable data labels, set each label to the point's exact value, and save the file as an .xlsx with Aspose.Cells for .NET.
class AddDataLabelsToColumnChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels for the series
            series.DataLabels.ShowValue = true;

            // Optionally set custom text for each point
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                point.DataLabels.Text = point.YValue.ToString();
            }

            // Save the workbook with the chart and data labels
            workbook.Save("ColumnChartWithDataLabels.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
