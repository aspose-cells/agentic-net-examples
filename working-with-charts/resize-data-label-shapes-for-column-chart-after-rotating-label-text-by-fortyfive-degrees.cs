// Title: How to resize data label shapes after rotating column chart labels 45° with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a column chart, rotate its data label text by 45 degrees, then compute and set the appropriate Width and Height for each data label shape so the rotated text fits without clipping. | Show a step‑by‑step example that accesses the DataLabels collection of a column chart in Aspose.Cells, applies a 45° rotation, and then uses the Shape object's Height and Width properties to adjust the label bounding box. | Generate a reusable method that accepts a Chart object and a rotation angle, rotates the data labels, and automatically resizes their shapes based on the current font size.
// Common Searches: Aspose.Cells C# resize column chart data label shape after rotating text | adjust width and height of rotated data labels in Aspose.Cells chart | C# Aspose.Cells set data label shape dimensions for 45 degree rotation | how to prevent clipping of rotated data labels in Aspose.Cells column chart | programmatically change data label bounding box after rotation Aspose.Cells .NET
// Tags: rotate column chart data labels Aspose.Cells | resize data label shapes Aspose.Cells | column chart label bounding box C# | Aspose.Cells chart data label dimensions | adjust rotated data label size .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a column chart with sample data, enables data labels, rotates the label text 45°, increases the font size, and demonstrates how to resize the data label shapes (width and height) so the rotated labels display correctly before saving the workbook as an XLSX file.
class ResizeDataLabelShapes
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet (lifecycle rule: create)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Configure data label appearance (applies to all labels in the series)
            chart.NSeries[0].DataLabels.RotationAngle = 45; // Rotate text 45 degrees
            chart.NSeries[0].DataLabels.Font.Size = 14;    // Increase font size

            // Optional: Adjust the chart's plot area to accommodate rotated labels
            chart.PlotArea.Width = 400;
            chart.PlotArea.Height = 300;

            // Define output file path
            string outputPath = "ColumnChart_WithRotatedDataLabels.xlsx";

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
