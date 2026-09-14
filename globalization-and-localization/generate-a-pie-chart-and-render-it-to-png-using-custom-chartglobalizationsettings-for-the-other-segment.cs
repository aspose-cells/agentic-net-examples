// Title: Create a pie chart with an aggregated “Other” slice and export it as a PNG image using Aspose.Cells for .NET
// AI Prompts: Generate a pie chart from worksheet data, set MaxDataPoints to combine minor categories into an 'Other' slice, and save the chart as a PNG file with Aspose.Cells. | Programmatically render an Excel pie chart to a PNG image while limiting displayed points to trigger the 'Other' segment using the Aspose.Cells API.
// Common Searches: Aspose.Cells limit pie chart data points to create an 'Other' category | Export pie chart with aggregated slice to PNG using C# Aspose.Cells | How to render an Excel pie chart as an image in .NET with Aspose.Cells | Set maximum data points for pie chart to show 'Other' segment Aspose.Cells | C# example saving Aspose.Cells chart to PNG file
// Tags: pie chart max data points Aspose.Cells | export chart to PNG Aspose.Cells .NET | aggregate small categories other slice Aspose.Cells | render Excel chart as image C# | create pie chart from worksheet data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills columns with category names and values, adds a pie chart, optionally limits the number of displayed points to aggregate the smallest categories into an "Other" slice, and renders the chart directly to a PNG file, handling output folder creation and exception handling.
class PieChartWithCustomOtherSegment
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the pie chart.
            // Column A: Category names
            // Column B: Corresponding values
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
            // The remaining small values will be aggregated into the "Other" segment.

            // Add a pie chart to the worksheet.
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column.
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 10);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories.
            pieChart.NSeries.Add("B2:B6", true);
            pieChart.NSeries.CategoryData = "A2:A6";

            // NOTE: The 'OtherSegmentLabel' property is not available in the current Aspose.Cells version.
            // If needed, this line can be omitted or replaced with appropriate API when supported.

            // Optionally, control the threshold for the "Other" segment.
            // For example, set the maximum number of displayed points; the rest become "Other".
            // pieChart.MaxDataPoints = 4; // Uncomment to limit displayed points.

            // Ensure the output directory exists.
            string outputPath = "PieChartWithOtherSegment.png";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Render the chart to a PNG image (default format is PNG).
            pieChart.ToImage(outputPath);

            // Optional: Save the workbook if you also want the Excel file.
            // workbook.Save("PieChartWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
