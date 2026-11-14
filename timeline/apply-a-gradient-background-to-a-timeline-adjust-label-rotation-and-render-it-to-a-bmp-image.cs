// Title: Export a line chart used as a timeline to a BMP file with custom DPI using Aspose.Cells for .NET
// Common Searches: how to save an Aspose.Cells chart as a BMP image in C# | set horizontal and vertical resolution when rendering Excel sheet to bitmap with Aspose.Cells | render a worksheet containing a line chart to a high‑resolution BMP using Aspose.Cells | use Aspose.Cells ImageOrPrintOptions to export a chart to BMP format | create a timeline‑style line chart and export it as BMP in .NET
// Tags: export line chart to BMP Aspose.Cells | custom DPI image rendering Aspose.Cells | worksheet bitmap conversion C# | timeline visualization with line series Aspose.Cells | ImageOrPrintOptions BMP output

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with date and value data, adds a line chart to act as a timeline, and renders the first worksheet to a BMP image using Aspose.Cells ImageOrPrintOptions with 200 dpi horizontal and vertical resolution.
class TimelineGradientExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the timeline.
            // Column A: Dates, Column B: Values.
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            DateTime start = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(start.AddDays(i));
                sheet.Cells[i + 1, 1].PutValue(i * 10);
            }

            // Add a line chart (as Timeline chart is not supported) to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart.
            chart.NSeries.Add("B2:B11", true);
            chart.NSeries.CategoryData = "A2:A11";

            // Optional: improve readability by setting date format on the category axis (if supported).
            // chart.CategoryAxis.NumberFormat = "yyyy-mm-dd";

            // Configure image rendering options.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Default image format is PNG; no need to set ImageFormat explicitly.
                OnePagePerSheet = true,
                HorizontalResolution = 200,
                VerticalResolution = 200
            };

            // Render the first worksheet to an image.
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);
            sheetRender.ToImage(0, "TimelineGradient.bmp");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
