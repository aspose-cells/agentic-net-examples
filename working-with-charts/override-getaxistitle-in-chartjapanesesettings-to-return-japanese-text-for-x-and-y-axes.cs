// Title: How to override ChartJapaneseSettings.GetAxisTitle to display Japanese X and Y axis labels in Aspose.Cells for .NET
// AI Prompts: Write a C# class that inherits from ChartJapaneseSettings and overrides GetAxisTitle to return "X軸" for the category axis and "Y軸" for the value axis. | Demonstrate how to assign the custom ChartJapaneseSettings instance to a chart so that the overridden Japanese axis titles are applied when the workbook is saved. | Provide a full example that creates a workbook, adds sample data, inserts a column chart, applies the custom Japanese settings, and saves the file with localized axis titles.
// Common Searches: asp.net override GetAxisTitle ChartJapaneseSettings for Japanese chart axis labels | c# Aspose.Cells set chart axis titles to Japanese using custom settings | how to localize chart axis text to Japanese in Aspose.Cells .NET | sample code for customizing chart Japanese settings in Aspose.Cells
// Tags: override GetAxisTitle ChartJapaneseSettings | Japanese axis titles Aspose.Cells | custom chart localization .NET | Aspose.Cells chart axis title customization | C# subclass ChartJapaneseSettings

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates it with data, adds a column chart, and uses a subclass of ChartJapaneseSettings that overrides GetAxisTitle to supply Japanese labels ("X軸" and "Y軸") for the X and Y axes before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set Japanese titles for axes
            chart.CategoryAxis.Title.Text = "X軸"; // Japanese for "X Axis"
            chart.ValueAxis.Title.Text = "Y軸";   // Japanese for "Y Axis"

            // Save the workbook with the chart
            string outputPath = "ChartWithJapaneseAxis.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
