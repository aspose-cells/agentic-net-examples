// Title: How to add a column chart in Aspose.Cells C# and use a separate column for custom category data labels
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells, sets the series values, and links the data labels to a cell range containing custom category names. | Show how to enable ShowCellRange and assign LinkedSource for a chart series to display custom labels from column C in an XLSX workbook. | Provide a complete example that formats the data labels (font color, position) and saves the workbook as ChartWithCustomCategoryLabels.xlsx.
// Common Searches: Aspose.Cells C# column chart custom data labels from another column | link cell range to chart series data labels Aspose.Cells .NET example | display category names as data labels in Aspose.Cells chart | set ShowCellRange and LinkedSource for chart labels using Aspose.Cells | how to format data label font color and position in Aspose.Cells chart
// Tags: Aspose.Cells column chart label linking | custom label range for chart series | ShowCellRange Aspose.Cells usage | LinkedSource property chart labels | chart data label formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsChartExample
{
    // The program creates a new workbook, fills columns A‑C with categories, values, and custom labels, adds a column chart, links the series data labels to the label range C2:C4 using ShowCellRange and LinkedSource, applies font color and position formatting, and saves the file as ChartWithCustomCategoryLabels.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data:
            // Column A – original categories (used for the chart axis)
            // Column B – numeric values for the series
            // Column C – custom labels that will be shown as data labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["A4"].PutValue("Gamma");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(45);

            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("First");
            sheet.Cells["C3"].PutValue("Second");
            sheet.Cells["C4"].PutValue("Third");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data (values) and category axis data (original categories)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels to use the custom label range (C2:C4)
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // optional: also show the numeric value
            series.DataLabels.ShowCellRange = true;      // enable using a cell range for labels
            series.DataLabels.LinkedSource = "C2:C4";    // link to the custom label column
            series.DataLabels.Font.Color = Color.DarkBlue;
            series.DataLabels.Position = LabelPositionType.InsideEnd;

            // Save the workbook
            workbook.Save("ChartWithCustomCategoryLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}
