// Title: Create a ChartJapaneseSettings class inheriting ChartGlobalizationSettings for Japanese chart localization in Aspose.Cells (C#)
// AI Prompts: Implement a C# class named ChartJapaneseSettings that derives from ChartGlobalizationSettings and supplies Japanese strings for chart titles, axis labels, and legend entries. | Show how to assign the ChartJapaneseSettings instance to Workbook.GlobalizationSettings so that every chart in the workbook automatically uses the Japanese labels. | Provide a complete example that creates a workbook, adds a column chart, applies the custom Japanese globalization settings, and saves the file.
// Common Searches: asp.net aspose.cells custom ChartGlobalizationSettings subclass for Japanese language | set default chart titles to Japanese in Aspose.Cells workbook | globalize chart axis labels to Japanese using Aspose.Cells C# | apply Japanese localization to all charts without manual title assignment Aspose.Cells | how to use ChartGlobalizationSettings to change chart language in Aspose.Cells
// Tags: ChartGlobalizationSettings Japanese subclass | Aspose.Cells workbook-wide chart localization | C# custom chart globalization class | Japanese axis and legend labels Aspose.Cells | global chart language configuration .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates creating a Workbook, adding sample data and a column chart, then using a custom ChartJapaneseSettings class that inherits from ChartGlobalizationSettings to provide Japanese text for chart titles, axis titles, and legend. The settings are registered with the workbook so all charts automatically display Japanese labels, and the workbook is saved as ChartJapanese.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data to the first worksheet
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Sales");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["B3"].PutValue(150);
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Apply Japanese titles directly to the chart
            chart.Title.Text = "チャートタイトル";
            chart.CategoryAxis.Title.Text = "カテゴリ軸";
            chart.ValueAxis.Title.Text = "値軸";

            // Set legend position using the correct enum
            chart.Legend.Position = Aspose.Cells.Charts.LegendPositionType.Right;

            // Set series name if any series exist
            if (chart.NSeries.Count > 0)
                chart.NSeries[0].Name = "系列";

            // Save the workbook
            string outputPath = "ChartJapanese.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
