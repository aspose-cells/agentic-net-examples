// Title: How to automatically set the Y‑axis display unit for a column chart in Aspose.Cells using C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and selects the appropriate ValueAxis.DisplayUnit based on the highest data point. | Show how to enable the unit label on the chart’s value axis so it displays "Millions", "Thousands", etc., mimicking Excel’s Auto option. | Write a reusable method that receives a worksheet range and returns the correct DisplayUnitType for the chart’s Y‑axis.
// Common Searches: Aspose.Cells C# automatically choose Y axis display unit for column chart | set chart value axis scaling to millions in Aspose.Cells | how to show unit label on Aspose.Cells chart axis | C# determine DisplayUnitType for Aspose chart based on data values | Excel-like auto display unit for Aspose.Cells column chart
// Tags: Aspose.Cells chart Y axis scaling | C# automatic Y axis unit selection Aspose chart | Aspose.Cells column chart axis unit label | Determine DisplayUnitType from max series value | Excel-like auto axis unit in Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAutoDisplayUnitDemo
{
    // The example creates a workbook, fills it with large numeric values, adds a column chart, computes the maximum series value, sets the chart's ValueAxis.DisplayUnit to the matching DisplayUnitType (e.g., Millions), enables the axis unit label, and saves the file as AutoDisplayUnitChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (large values to demonstrate automatic unit selection)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1500000);   // 1.5 million
            sheet.Cells["B3"].PutValue(3000000);   // 3 million
            sheet.Cells["B4"].PutValue(4500000);   // 4.5 million

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Determine the maximum value in the series to decide which display unit to use
            double maxValue = Math.Max(Math.Max(sheet.Cells["B2"].DoubleValue,
                                                sheet.Cells["B3"].DoubleValue),
                                                sheet.Cells["B4"].DoubleValue);

            // Apply an automatic display unit similar to Excel's "Auto" option
            if (maxValue >= 1_000_000_000)               // Billions
                chart.ValueAxis.DisplayUnit = DisplayUnitType.Billions;
            else if (maxValue >= 100_000_000)            // Hundred Millions
                chart.ValueAxis.DisplayUnit = DisplayUnitType.HundredMillions;
            else if (maxValue >= 10_000_000)             // Ten Millions
                chart.ValueAxis.DisplayUnit = DisplayUnitType.TenMillions;
            else if (maxValue >= 1_000_000)              // Millions
                chart.ValueAxis.DisplayUnit = DisplayUnitType.Millions;
            else if (maxValue >= 100_000)                // Hundred Thousands
                chart.ValueAxis.DisplayUnit = DisplayUnitType.HundredThousands;
            else if (maxValue >= 10_000)                 // Ten Thousands
                chart.ValueAxis.DisplayUnit = DisplayUnitType.TenThousands;
            else if (maxValue >= 1_000)                  // Thousands
                chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            else
                chart.ValueAxis.DisplayUnit = DisplayUnitType.None; // No scaling needed

            // Optionally show the unit label (e.g., "Millions")
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Save the workbook
            workbook.Save("AutoDisplayUnitChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
