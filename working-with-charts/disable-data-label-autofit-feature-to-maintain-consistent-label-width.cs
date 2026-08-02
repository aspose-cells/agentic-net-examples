// Title: Disable data label auto‑fit and set fixed width in Aspose.Cells chart (C#)
// Description: Creates a workbook, adds a column chart, enables data labels, then turns off automatic shape resizing and assigns a 60‑pixel width to each label before saving the file.
// Keywords: Aspose.Cells C# chart data labels | disable auto‑fit data labels | fixed label width | IsResizeShapeToFitText false | WidthPixel property | column chart labeling
// Common Searches: How to prevent chart data labels from resizing in Aspose.Cells | Set constant pixel width for Excel chart labels using .NET | Aspose.Cells chart label size control | C# code to fix data label width in column chart
// Developer Intent: Turn off automatic label resizing and apply a uniform pixel width to chart data labels.
// Use Cases: Financial dashboards that require aligned label columns | Standardized Excel reports where label dimensions must stay constant | Generating PDFs or images from charts with consistent label sizing
// AI Prompts: Provide C# code to stop data label auto‑fit in an Aspose.Cells chart and set each label to 60 pixels wide. | Show how to apply a fixed label width to multiple series in a workbook using Aspose.Cells. | Explain which properties keep chart data label dimensions constant in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelAutoFitDemo
{
    // Creates a workbook, adds a column chart, enables data labels, then turns off automatic shape resizing and assigns a 60‑pixel width to each label before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Disable auto‑fit for each data label and set a fixed width
            foreach (ChartPoint point in series.Points)
            {
                // Prevent the label shape from resizing to fit its text
                point.DataLabels.IsResizeShapeToFitText = false;

                // Set a consistent width (in pixels) for all labels
                point.DataLabels.WidthPixel = 60;
            }

            // Save the workbook
            workbook.Save("DataLabelAutoFitDisabled.xlsx");
        }
    }
}
