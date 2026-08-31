// Title: Set a custom font for Excel chart data labels and auto‑resize label shapes with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a column chart, turns on data labels, assigns Calibri 14 bold dark‑blue font to the labels, and enables the label shapes to resize automatically to fit the text using Aspose.Cells. | Update an existing Aspose.Cells chart series so that its data label font is bold Calibri 14 dark blue and the labels automatically adjust their shape size and font scaling.
// Common Searches: asp.net how to change font of chart data labels with Aspose.Cells | asp.net chart data label shape auto resize Aspose.Cells example | c# set bold dark blue font for Excel chart data labels using Aspose.Cells | asp.net enable autoscale font for chart data labels Aspose.Cells | apply custom font to all data label nodes in Aspose.Cells chart series
// Tags: set chart data label font Aspose.Cells | enable data label shape auto‑resize Aspose.Cells | auto‑scale chart data label font .NET | custom font for series data labels Aspose.Cells | column chart label styling Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomDataLabelFont
{
    // The example creates a workbook, adds sample data, inserts a column chart, enables data labels for the first series, applies a Calibri 14 bold dark‑blue font to the labels, activates shape resizing and auto‑scaling for the labels, propagates the font to child label nodes, and saves the file as CustomDataLabelFont.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply custom font style to the data labels
            series.DataLabels.Font.Name = "Calibri";
            series.DataLabels.Font.Size = 14;
            series.DataLabels.Font.Color = Color.DarkBlue;
            series.DataLabels.Font.IsBold = true;

            // Ensure the data label shape resizes to fit the new font
            series.DataLabels.IsResizeShapeToFitText = true;
            // Optionally enable auto‑scaling of the font when the shape size changes
            series.DataLabels.AutoScaleFont = true;

            // Apply the font settings to all child label nodes
            series.DataLabels.ApplyFont();

            // Save the workbook (save rule)
            workbook.Save("CustomDataLabelFont.xlsx");
        }
    }
}
