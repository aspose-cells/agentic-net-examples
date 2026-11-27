// Title: How to resize chart data label shapes after applying superscript to part of the label text with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# resize data label after setting superscript on characters | set partial superscript in chart data label and change label dimensions Aspose.Cells | disable automatic label shape resizing in Aspose.Cells chart and specify height and width | adjust chart data label size for superscript text using Aspose.Cells .NET | customize data label shape size for column chart in Aspose.Cells after formatting
// Tags: chart data label shape resizing Aspose.Cells | partial superscript in chart label .NET | manual data label height width Aspose.Cells | column chart label size customization C# | disable auto label resize Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace ResizeDataLabelAfterSuperscript
{
    // The example creates a workbook, adds a column chart, enables data labels, applies superscript formatting to the first two characters of each label, disables automatic shape resizing, and manually sets the label height and width to accommodate the superscript before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["B4"].PutValue(45);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels seriesLabels = chart.NSeries[0].DataLabels;
            seriesLabels.ShowValue = true;
            seriesLabels.Position = LabelPositionType.InsideEnd;

            // Apply superscript style to the first two characters of each label
            // and then manually resize the label shape to accommodate the new formatting
            foreach (ChartPoint point in chart.NSeries[0].Points)
            {
                // Access the DataLabels of the individual point
                DataLabels pointLabel = point.DataLabels;

                // Ensure the label is visible
                pointLabel.ShowValue = true;

                // Apply superscript to the first two characters (e.g., the numeric part)
                // Characters method is 0‑based; length 2 means characters 0 and 1
                var chars = pointLabel.Characters(0, 2);
                chars.Font.IsSuperscript = true;

                // Disable automatic resizing so we can control the shape size
                pointLabel.IsResizeShapeToFitText = false;

                // Manually increase the height to give room for superscript text
                // Height is measured in points (1/72 inch). Adjust as needed.
                pointLabel.Height = 30; // default is ~15, increase to 30
                pointLabel.Width = 60;  // ensure width is sufficient

                // Optionally, you can also adjust the font size to keep the label readable
                pointLabel.Font.Size = 10;
            }

            // Save the workbook
            string outputPath = "ResizeDataLabelSuperscript.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
