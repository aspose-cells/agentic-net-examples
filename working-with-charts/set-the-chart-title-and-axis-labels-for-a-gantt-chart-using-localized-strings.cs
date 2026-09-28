// Title: Set a localized chart title and axis labels for a Gantt chart using Aspose.Cells in C#
// AI Prompts: Generate a Gantt chart with Aspose.Cells and assign French‑localized strings to the chart title, X‑axis, and Y‑axis. | Implement a GetLocalizedString method that retrieves resource values based on CultureInfo and use it to set chart titles for multiple locales in Aspose.Cells. | Modify the example to use a German CultureInfo (de-DE) and update the chart title and axis labels accordingly.
// Common Searches: aspnet how to apply culture-specific titles to charts created with Aspose.Cells | c# set chart title and axis labels for a Gantt chart using Aspose.Cells | localize Aspose.Cells chart axis text for French locale | example of stacked bar Gantt chart with localized strings in Aspose.Cells C# | retrieve resource strings for chart titles based on CultureInfo in Aspose.Cells
// Tags: Aspose.Cells chart title localization | Aspose.Cells Gantt chart implementation | Aspose.Cells axis title cultureinfo | resource-based localization for Aspose.Cells charts | set chart titles with CultureInfo in Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Globalization;

// Demonstrates creating a stacked‑bar Gantt chart in a new workbook, retrieving French‑localized strings via a mock GetLocalizedString method, assigning them to the chart title, X‑axis, and Y‑axis, and saving the workbook as GanttChart.xlsx.
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

            // Sample data for a Gantt chart
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Start");
            sheet.Cells["C1"].PutValue("Duration");

            sheet.Cells["A2"].PutValue("Design");
            sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Development");
            sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 6));
            sheet.Cells["C3"].PutValue(10);

            sheet.Cells["A4"].PutValue("Testing");
            sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 16));
            sheet.Cells["C4"].PutValue(4);

            // Add a stacked bar chart (used to represent a Gantt chart)
            // Correct enum name is BarStacked
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Localized strings (example for French locale)
            CultureInfo ci = new CultureInfo("fr-FR");
            string chartTitle = GetLocalizedString("ChartTitle", ci);   // "Plan du projet"
            string xAxisLabel = GetLocalizedString("XAxisLabel", ci);   // "Date de début"
            string yAxisLabel = GetLocalizedString("YAxisLabel", ci);   // "Tâches"

            // Set chart title
            chart.Title.Text = chartTitle;
            chart.Title.Font.IsBold = true;

            // Set axis titles
            chart.ValueAxis.Title.Text = xAxisLabel;      // X‑axis (value axis)
            chart.CategoryAxis.Title.Text = yAxisLabel;   // Y‑axis (category axis)

            // Save the workbook
            workbook.Save("GanttChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple mock localization method; replace with real resource lookup as needed
    static string GetLocalizedString(string key, CultureInfo culture)
    {
        if (culture.Name == "fr-FR")
        {
            return key switch
            {
                "ChartTitle" => "Plan du projet",
                "XAxisLabel" => "Date de début",
                "YAxisLabel" => "Tâches",
                _ => key
            };
        }

        // Default English strings
        return key switch
        {
            "ChartTitle" => "Project Plan",
            "XAxisLabel" => "Start Date",
            "YAxisLabel" => "Tasks",
            _ => key
        };
    }
}
