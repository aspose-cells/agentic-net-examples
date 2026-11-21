// Title: Set Japanese X‑axis and Y‑axis titles on a column chart using Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells set chart axis title to Japanese characters | aspocells column chart axis label localization example | how to add Japanese text to CategoryAxis.Title in Aspose.Cells | assign Unicode Japanese text to ValueAxis.Title with Aspose.Cells | create Excel chart with Japanese axis titles using Aspose.Cells .NET
// Tags: Aspose.Cells column chart axis title localization | C# set CategoryAxis.Title Japanese text | Aspose.Cells ValueAxis.Title Unicode characters | Excel workbook chart axis labels Japanese Aspose.Cells | Aspose.Cells chart axis title assignment C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsJapaneseAxisTitle
{
    // Demonstrates creating a workbook, adding sample data, inserting a column chart, and assigning Japanese strings to the CategoryAxis and ValueAxis titles using Aspose.Cells for .NET, then saving the file as an .xlsx workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart.
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart data source.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set Japanese titles for the axes.
                chart.CategoryAxis.Title.Text = "カテゴリ"; // X‑axis (category) title
                chart.ValueAxis.Title.Text = "値";          // Y‑axis (value) title

                // Save the workbook to a file.
                string outputPath = "ChartWithJapaneseAxisTitles.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
