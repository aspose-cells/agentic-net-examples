// Title: C# – Toggle Visibility of Individual Waterfall Chart Series with Aspose.Cells
// Description: Creates a workbook, adds two data series to a Waterfall chart, and uses a boolean array to set each series' IsFiltered flag, enabling programmatic show/hide of chart series before saving the Excel file.
// Keywords: Aspose.Cells C# | Waterfall chart series visibility | IsFiltered property | hide chart series .NET | show chart series programmatically | dynamic chart filtering | Excel export Aspose.Cells | GitHub Aspose.Cells example | chart series toggle code | Aspose.Cells chart API
// Common Searches: how to hide a series in Aspose.Cells Waterfall chart | set chart series visibility based on user input Aspose.Cells | use IsFiltered to filter chart series C# | toggle multiple Waterfall series programmatically | Aspose.Cells example for dynamic chart series
// Developer Intent: The developer needs to control which Waterfall chart series are displayed by programmatically setting their visibility according to runtime preferences.
// Use Cases: Interactive dashboards where users select which series to display before exporting to Excel. | Financial reports that only show user‑chosen data streams in a Waterfall chart. | Automated scripts that remove irrelevant series to keep charts concise during batch processing.
// AI Prompts: Generate C# code using Aspose.Cells that toggles any number of chart series based on a boolean array. | Explain the effect of the IsFiltered property on chart rendering and how to use it for visibility control. | Show how to bind WinForms checkboxes to Waterfall chart series visibility with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallSeriesVisibilityDemo
{
    // Creates a workbook, adds two data series to a Waterfall chart, and uses a boolean array to set each series' IsFiltered flag, enabling programmatic show/hide of chart series before saving the Excel file.
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

            // Values column (multiple series to demonstrate toggling)
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(-20);
            sheet.Cells["B5"].PutValue(110);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(-15);
            sheet.Cells["C5"].PutValue(90);

            // Add a Waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add both series to the chart
            chart.NSeries.Add("B2:B5", true); // Series 1
            chart.NSeries.Add("C2:C5", true); // Series 2
            chart.NSeries.CategoryData = "A2:A5";

            // Simulated user preferences: true = show series, false = hide series
            // For example, hide Series 2 and show Series 1
            bool[] userPreferences = new bool[] { true, false };

            // Apply visibility based on preferences
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                // If a series is filtered (IsFiltered = true) it will NOT be displayed.
                // Therefore, set IsFiltered to the opposite of the user's "show" preference.
                chart.NSeries[i].IsFiltered = !userPreferences[i];
            }

            // Save the workbook
            workbook.Save("WaterfallSeriesVisibilityDemo.xlsx");
        }
    }
}
