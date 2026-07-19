// Title: Toggle Waterfall Chart Series Visibility with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data and a Waterfall chart, then uses a Dictionary<int, bool> to set each series' IsFiltered flag, hiding or showing series according to user preferences before saving the file.
// Keywords: Aspose.Cells C# Waterfall chart | chart series visibility | IsFiltered property | hide chart series programmatically | dynamic chart filtering Aspose | Excel workbook series toggle | user‑driven chart customization
// Common Searches: Aspose.Cells hide waterfall series C# | set IsFiltered for chart series .NET | programmatically filter chart series Aspose | toggle visibility of Excel chart series using code | waterfall chart series visibility dictionary
// Developer Intent: Control the display of individual Waterfall chart series based on runtime settings.
// Use Cases: Hide secondary data series in a financial waterfall report while keeping primary values visible. | Apply checkbox selections from a UI to show or hide specific chart series in generated workbooks. | Create a reusable routine that respects user‑defined visibility flags for any chart type.
// AI Prompts: Write a C# method that receives a Chart object and a Dictionary<int, bool> and updates each series' IsFiltered property accordingly. | Explain how the IsFiltered flag influences series rendering in Aspose.Cells charts. | Generate a generic helper to toggle visibility for NSeries across different chart types using Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallSeriesVisibilityDemo
{
    // Creates a workbook, adds sample data and a Waterfall chart, then uses a Dictionary<int, bool> to set each series' IsFiltered flag, hiding or showing series according to user preferences before saving the file.
    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook and worksheet --------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];               // get the first worksheet

            // -------------------- Populate sample data for a Waterfall chart --------------------
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["A5"].PutValue("Profit");

            // Series 1 values
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(-50);
            sheet.Cells["B5"].PutValue(100);

            // Series 2 values
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(120);
            sheet.Cells["C3"].PutValue(130);
            sheet.Cells["C4"].PutValue(-30);
            sheet.Cells["C5"].PutValue(100);

            // -------------------- Add a Waterfall chart --------------------
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add both series to the chart
            chart.NSeries.Add("B2:B5", true);   // Series 1
            chart.NSeries.Add("C2:C5", true);   // Series 2
            chart.NSeries.CategoryData = "A2:A5";

            // -------------------- User preferences for series visibility --------------------
            // Example: user wants to hide Series 2 but keep Series 1 visible
            // The key is the zero‑based series index, the value indicates whether the series should be shown.
            Dictionary<int, bool> userPreferences = new Dictionary<int, bool>
            {
                { 0, true },   // Series 0 (first series) -> visible
                { 1, false }   // Series 1 (second series) -> hidden
            };

            // -------------------- Apply visibility using IsFiltered property --------------------
            // IsFiltered = true  => series is filtered (hidden)
            // IsFiltered = false => series is displayed
            foreach (KeyValuePair<int, bool> kvp in userPreferences)
            {
                int seriesIndex = kvp.Key;
                bool shouldShow = kvp.Value;

                if (seriesIndex >= 0 && seriesIndex < chart.NSeries.Count)
                {
                    chart.NSeries[seriesIndex].IsFiltered = !shouldShow;
                }
            }

            // -------------------- Save the workbook --------------------
            workbook.Save("WaterfallSeriesVisibilityDemo.xlsx");
        }
    }
}
