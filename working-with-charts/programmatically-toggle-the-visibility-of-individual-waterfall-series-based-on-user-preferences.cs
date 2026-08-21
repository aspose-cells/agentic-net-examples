// Title: Toggle Waterfall chart series visibility programmatically with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data and a Waterfall chart with two series, then uses a boolean array to set each series' IsFiltered property, showing or hiding the series before saving the file.
// Keywords: Aspose.Cells | Waterfall chart | series visibility | IsFiltered | .NET | C# | toggle chart series | hide Excel series | programmatic chart filter | Excel automation
// Common Searches: Aspose.Cells hide waterfall series | How to filter chart series in Aspose.Cells | Set IsFiltered property C# | Toggle chart series visibility .NET | Programmatically show or hide Excel chart series
// Developer Intent: Show or hide individual Waterfall chart series based on user‑defined preferences.
// Use Cases: Allow users to deselect secondary series in a financial waterfall report via UI controls. | Create interactive Excel dashboards where checkboxes toggle series before export. | Generate scenario‑specific waterfall charts by displaying only the relevant series.
// AI Prompts: Write C# code that reads series visibility flags from a JSON file and applies them to a Waterfall chart using Aspose.Cells' IsFiltered property. | Show how to bind WinForms checkboxes to each Waterfall series' IsFiltered setting for real‑time toggling. | Explain how to refresh a Waterfall chart after changing IsFiltered values so the changes appear in the saved workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallSeriesVisibilityDemo
{
    // Creates a workbook, adds sample data and a Waterfall chart with two series, then uses a boolean array to set each series' IsFiltered property, showing or hiding the series before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a Waterfall chart
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["A5"].PutValue("End");

            // Values column (multiple series for demonstration)
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(-20);
            sheet.Cells["B5"].PutValue(110);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(120);
            sheet.Cells["C3"].PutValue(40);
            sheet.Cells["C4"].PutValue(-10);
            sheet.Cells["C5"].PutValue(150);

            // Add a Waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Simulated user preferences: true = visible, false = hidden
            bool[] userPreferences = new bool[] { true, false }; // Series1 visible, Series2 hidden

            // Apply visibility based on preferences using IsFiltered property
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                // If a series is filtered (IsFiltered = true) it will NOT be displayed.
                // Therefore we set IsFiltered to the inverse of the user's visibility choice.
                chart.NSeries[i].IsFiltered = !userPreferences[i];
            }

            // Save the workbook
            workbook.Save("WaterfallSeriesVisibilityDemo.xlsx");
        }
    }
}
