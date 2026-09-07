// Title: How to set column chart series data labels to InsideEnd position using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook with a column chart and configure the first series to display data labels positioned at the InsideEnd of each column using Aspose.Cells. | Modify an existing Aspose.Cells column chart so that its series data labels are shown inside the column ends for improved readability. | Generate a .xlsx file where a column chart automatically shows values inside the column tops by setting LabelPositionType.InsideEnd on the series data labels.
// Common Searches: Aspose.Cells C# set column chart data label position to InsideEnd | How to display data labels inside the end of columns in an Excel chart using Aspose.Cells | C# Aspose.Cells example for positioning series data labels inside column ends | Set label position InsideEnd for column chart series with Aspose.Cells .NET
// Tags: Aspose.Cells column chart data label positioning | C# set label position InsideEnd | Aspose.Cells series data labels inside end | Excel column chart label placement .NET | Aspose.Cells chart customization C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDataLabelPosition
{
    // // This C# program creates a workbook, adds sample data, inserts a column chart, enables data labels for the first series, sets the label position to InsideEnd, and saves the file as ColumnChart_With_InsideEndDataLabels.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
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

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Set data label position to InsideEnd for better readability
            series.DataLabels.Position = LabelPositionType.InsideEnd;

            // Save the workbook to a file
            workbook.Save("ColumnChart_With_InsideEndDataLabels.xlsx");
        }
    }
}
