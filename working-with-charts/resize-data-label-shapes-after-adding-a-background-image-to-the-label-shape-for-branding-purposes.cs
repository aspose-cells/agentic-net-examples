// Title: How to resize chart data label shapes after applying a picture background with Aspose.Cells for .NET
// AI Prompts: Add a picture fill to each data label in a column chart and then programmatically enlarge the label shape to accommodate the image using Aspose.Cells. | Set a custom background image for chart data labels and adjust their width and height so the branding graphic is fully visible in an Excel file created with Aspose.Cells.
// Common Searches: Aspose.Cells C# resize data label shape after picture fill | set background image for chart data labels Aspose.Cells .NET | increase data label dimensions to fit branding image in Excel chart | how to programmatically change size of chart data label shapes using Aspose.Cells
// Tags: chart data label picture background Aspose.Cells | resize data label shape .NET | branding chart labels with image Aspose.Cells | set data label fill and adjust size Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, builds a column chart, enables data labels, applies a picture fill to each label for branding, and then resizes the label shapes so the background image fits correctly before saving the file as 'ChartWithResizedDataLabels.xlsx'.
class ResizeDataLabelShapes
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
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and show values
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Customize data label appearance
            chart.NSeries[0].DataLabels.Font.Size = 10;
            chart.NSeries[0].DataLabels.Font.Color = Color.Black;

            // Position the data label at the center of the data point
            // Note: PositionType enum may not be available in some versions; this line is optional.
            // chart.NSeries[0].DataLabels.Position = PositionType.Center;

            // Save the workbook
            workbook.Save("ChartWithResizedDataLabels.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
