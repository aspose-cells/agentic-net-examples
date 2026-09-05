// Title: Enable cell‑based data labels for a column chart series with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets Series.DataLabels.ShowCellRange to true and links the labels to a cell range using Series.DataLabels.LinkedSource in an Aspose.Cells chart. | Show how to display the numeric value, apply a custom font color, and use cell‑based labels for a column chart series in Aspose.Cells.
// Common Searches: Aspose.Cells C# show data labels from cells in a column chart | How to link custom label text to a chart series using Aspose.Cells | Set ShowCellRange property for series data labels in Aspose.Cells .NET | Display cell range values as data labels in an Aspose.Cells chart | Change font color of data labels when using cell‑based labels in Aspose.Cells
// Tags: Series.DataLabels.ShowCellRange Aspose.Cells | link data label cells Aspose.Cells chart | column chart cell‑based data labels .NET | custom label font styling Aspose.Cells | chart series data label configuration C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsShowCellRangeDemo
{
    // The example creates a workbook, adds sample data, inserts a column chart, and configures the first series to show data labels sourced from cells C2:C3 by enabling ShowCellRange, setting LinkedSource, displaying the numeric value, and applying a blue font color, then saves the file as ShowCellRangeDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["C2"].PutValue("100 units");
            sheet.Cells["C3"].PutValue("200 units");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the series and categories
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels and activate cell‑based data labels
            series.DataLabels.ShowValue = true;          // optional: show the numeric value
            series.DataLabels.ShowCellRange = true;      // activate cell range as data labels
            series.DataLabels.LinkedSource = "C2:C3";    // link to cells containing custom label text
            series.DataLabels.Font.Color = Color.Blue;  // optional styling

            // Save the workbook
            workbook.Save("ShowCellRangeDemo.xlsx");
        }
    }
}
