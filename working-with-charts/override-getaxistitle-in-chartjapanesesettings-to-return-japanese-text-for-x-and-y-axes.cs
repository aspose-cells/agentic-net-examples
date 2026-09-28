// Title: How to set Japanese X‑axis and Y‑axis titles on a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook, adds a column chart, and assigns Japanese strings to the CategoryAxis and ValueAxis titles with Aspose.Cells. | Show how to enable visibility and customize the axis titles of a chart to display Japanese characters in an Aspose.Cells‑generated .xlsx file. | Provide a complete Aspose.Cells example that demonstrates localization of chart axis labels to Japanese and saves the workbook.
// Common Searches: Aspose.Cells C# set chart axis title to Japanese characters | How to display non‑Latin text on Excel chart axes with Aspose.Cells | C# example for adding Japanese labels to column chart axes in .xlsx | Localization of chart axis titles using Aspose.Cells for .NET
// Tags: Aspose.Cells chart axis title localization | C# set chart axis title Unicode | Aspose.Cells column chart Japanese labels | Excel chart axis title visibility Aspose.Cells | Aspose.Cells workbook save with localized chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartJapaneseExample
{
    // The sample creates a new Workbook, fills it with data, adds a column chart, sets Japanese text for both the Category (X) and Value (Y) axis titles, makes the titles visible, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure Japanese axis titles directly
                chart.CategoryAxis.Title.Text = "横軸タイトル"; // X Axis Title in Japanese
                chart.CategoryAxis.Title.IsVisible = true;
                chart.ValueAxis.Title.Text = "縦軸タイトル";   // Y Axis Title in Japanese
                chart.ValueAxis.Title.IsVisible = true;

                // Define output file path
                string outputPath = "ChartWithJapaneseAxisTitles.xlsx";

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
}
