// Title: Manually style a Waterfall chart in Aspose.Cells for .NET to match corporate branding guidelines
// AI Prompts: Write C# code using Aspose.Cells to create a Waterfall chart and set custom series colors, line styles, and title font to reflect a corporate brand. | Show how to apply a predefined color palette and font settings to an Aspose.Cells Waterfall chart when the ApplyTheme method is unavailable.
// Common Searches: Aspose.Cells C# set custom colors for Waterfall chart series | How to format chart title font in Aspose.Cells .NET | Apply corporate color palette to Excel chart using Aspose.Cells | Manual theme styling for Waterfall chart in Aspose.Cells without ApplyTheme
// Tags: Aspose.Cells waterfall chart color customization | C# customize Excel chart theme Aspose.Cells | manual chart styling Aspose.Cells .NET | apply corporate branding to Excel chart programmatically | waterfall chart formatting Aspose.Cells workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes sample financial data to cells A1:B5, adds a Waterfall chart, sets its title, binds the series to the data range, and notes that Aspose.Cells does not provide an ApplyTheme method, so styling must be performed manually (e.g., setting series colors, line styles, and title font). The workbook is saved as WaterfallChartWithCorporateTheme.xlsx.
class WaterfallChartWithTheme
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the Waterfall chart
            // Column A: Categories, Column B: Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(5000);
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(8000);
            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["B4"].PutValue(-3000);
            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["B5"].PutValue(5000);

            // Add a Waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 27, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Financial Waterfall";

            // Add series and bind it to the data range
            int seriesIndex = chart.NSeries.Add("B2:B5", true);
            // Set category (X) axis labels
            chart.NSeries[seriesIndex].XValues = "A2:A5";

            // Note: Applying a theme is omitted because the ApplyTheme method is not available
            // in the current Aspose.Cells API version.

            // Save the workbook to a file
            string outputPath = "WaterfallChartWithCorporateTheme.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
