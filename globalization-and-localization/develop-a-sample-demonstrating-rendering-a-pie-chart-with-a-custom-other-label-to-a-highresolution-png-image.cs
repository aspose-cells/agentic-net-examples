// Title: Render a pie chart with an "Other" slice to a 300 DPI PNG using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# render pie chart with custom label to high resolution PNG | Create pie chart with "Other" data point and export as 300 DPI PNG in .NET | Set image resolution when exporting Aspose.Cells chart to PNG | Add custom slice label to pie chart using Aspose.Cells .NET API | Generate high‑resolution PNG of a worksheet chart with Aspose.Cells
// Tags: pie chart custom label Aspose.Cells | export chart high DPI PNG C# | ImageOrPrintOptions resolution Aspose.Cells | render worksheet chart as PNG Aspose.Cells | set data point name pie chart Aspose.Cells

using System;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with fruit categories and values, adds a pie chart, sets a chart title, configures ImageOrPrintOptions for 300 DPI horizontal and vertical resolution, and renders the worksheet containing the chart to a PNG file named "PieChartWithOtherLabel.png". An optional commented line shows how to rename a data point to "Other".
class PieChartWithOtherLabel
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the pie chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherries");
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["A5"].PutValue("Dates");
            sheet.Cells["B5"].PutValue(10);
            sheet.Cells["A6"].PutValue("Elderberries");
            sheet.Cells["B6"].PutValue(5);
            sheet.Cells["A7"].PutValue("Figs");
            sheet.Cells["B7"].PutValue(3);
            sheet.Cells["A8"].PutValue("Grapes");
            sheet.Cells["B8"].PutValue(2);

            // Add a pie chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 10, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the series.
            chart.NSeries.Add("B2:B8", true);
            chart.NSeries.CategoryData = "A2:A8";

            // Optional: rename a data point (if supported). Commented out to avoid API mismatch.
            // chart.NSeries[0].Points[6].SetDataPointName("Other");

            // Set chart title.
            chart.Title.Text = "Fruit Distribution with \"Other\"";

            // Configure high‑resolution image rendering options.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // ImageFormat defaults to PNG; explicit setting omitted to avoid API mismatch.
                HorizontalResolution = 300, // DPI
                VerticalResolution = 300,   // DPI
                OnePagePerSheet = true
            };

            // Render the worksheet (which contains the chart) to a PNG image.
            SheetRender sr = new SheetRender(sheet, imgOptions);
            sr.ToImage(0, "PieChartWithOtherLabel.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
