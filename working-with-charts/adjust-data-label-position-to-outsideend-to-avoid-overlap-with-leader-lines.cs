// Title: Set column chart data labels to OutsideEnd and style leader lines with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a column chart, and sets the series data labels to the OutsideEnd position using Aspose.Cells. | Show how to enable leader lines for a chart series and customize their color, weight, and dash style in Aspose.Cells for .NET. | Provide a complete example that positions data labels outside the columns to avoid overlap and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# column chart label OutsideEnd to prevent overlap | How to add and format leader lines for a series in Aspose.Cells .NET | Example of positioning data labels outside columns in an Aspose.Cells chart | Styling chart leader lines color and weight with Aspose.Cells for C#
// Tags: Aspose.Cells set data label position OutsideEnd | Aspose.Cells customize chart leader line style | Aspose.Cells column chart label placement C# | Aspose.Cells avoid data label overlap .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsDataLabelPositionDemo
{
    // Demonstrates creating a workbook, adding a column chart, showing values in data labels, positioning those labels OutsideEnd to avoid overlap, enabling leader lines, customizing their appearance, and saving the result as DataLabelOutsideEndDemo.xlsx.
    class Program
    {
        static void Main()
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Position data labels outside the end of each column to avoid overlap with leader lines
            series.DataLabels.Position = LabelPositionType.OutsideEnd;

            // Enable leader lines (optional, demonstrates the effect)
            series.HasLeaderLines = true;
            // Customize leader line appearance
            series.LeaderLines.IsAuto = false;
            series.LeaderLines.Style = LineType.Solid;
            series.LeaderLines.WeightPt = 1.0;
            series.LeaderLines.Color = Color.Gray;

            // Save the workbook
            workbook.Save("DataLabelOutsideEndDemo.xlsx");
        }
    }
}
