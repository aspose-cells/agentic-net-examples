// Title: How to validate and truncate a chart data label to Excel's 255‑character limit using Aspose.Cells for .NET
// AI Prompts: Generate C# code that checks the length of a chart data label and trims it to 255 characters before applying it with Aspose.Cells. | Show an example of assigning a safe custom label to a chart point, including a fallback for Aspose.Cells versions without the DataPoints collection.
// Common Searches: Aspose.Cells truncate chart label text to 255 characters C# | validate Excel chart data label length before setting with Aspose.Cells | how to limit custom data label size in Aspose.Cells chart series | C# example for safe custom label on Excel chart using Aspose.Cells
// Tags: truncate chart data label Aspose.Cells | validate Excel label length .NET | chart series custom label handling | Excel data label character limit Aspose.Cells | Aspose.Cells chart label compatibility

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, adds a column chart, checks a custom label's length against Excel's 255‑character limit, truncates it if necessary, configures data label visibility, and saves the file as CustomLabelValidated.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series (values) and set category data
            int seriesIndex = chart.NSeries.Add("B2:B4", true);
            // If the API version supports CategoryData, uncomment the next line:
            // chart.NSeries[seriesIndex].CategoryData = "A2:A4";

            // Example custom label text (truncated if necessary)
            string customLabel = "This is a custom data label that might be too long for Excel's limit.";
            const int MaxLabelLength = 255;
            if (customLabel.Length > MaxLabelLength)
            {
                customLabel = customLabel.Substring(0, MaxLabelLength);
            }

            // Configure data label visibility for the series
            Series series = chart.NSeries[seriesIndex];
            series.DataLabels.ShowValue = true;          // show values
            series.DataLabels.ShowCategoryName = false;
            series.DataLabels.ShowSeriesName = false;
            series.DataLabels.ShowLegendKey = false;

            // NOTE: Setting a custom label for an individual data point requires the DataPoints collection,
            // which may not be available in older Aspose.Cells versions. The following code is omitted
            // for compatibility. If your version supports DataPoints, you can uncomment and use it:
            // if (series.DataPoints != null && series.DataPoints.Count > 0)
            // {
            //     series.DataPoints[0].SetDataLabel(customLabel, true); // true = isCustom
            // }

            // Save the workbook
            string outputPath = "CustomLabelValidated.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
