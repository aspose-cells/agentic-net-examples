// Title: How to toggle visibility of individual series in a Waterfall chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a dictionary of series indexes and boolean flags, then sets each corresponding series in an Aspose.Cells Waterfall chart to visible or hidden. | Show how to safely handle the absence of a series visibility property in Aspose.Cells by using a conditional placeholder in the loop. | Demonstrate saving the workbook after modifying chart series display states with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells hide specific series in a waterfall chart C# | programmatically change series visibility in Excel chart using Aspose.Cells .NET | toggle waterfall chart series on and off based on user settings Aspose.Cells | C# dictionary to control chart series display Aspose.Cells | IsVisible property for chart series in Aspose.Cells not available
// Tags: waterfall chart series display Aspose.Cells | C# Aspose.Cells chart series toggle | Aspose.Cells series display control | Excel workbook chart series hide show .NET | user-driven chart series management Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data and a Waterfall chart, defines a dictionary mapping series indexes to visibility flags, iterates through the chart's series to apply those flags (with a placeholder for the future IsVisible property), and saves the workbook as WaterfallToggleVisibility.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data for a Waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["B4"].PutValue(-20);
            sheet.Cells["A5"].PutValue("End");
            sheet.Cells["B5"].PutValue(110);

            // Add a Waterfall chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            var chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";     // Categories

            // Example user preferences: series index -> visibility flag
            var visibilityPreferences = new Dictionary<int, bool>()
            {
                { 0, true },   // Show first series
                { 1, false },  // Hide second series
                { 2, true },   // Show third series
                { 3, false }   // Hide fourth series
            };

            // Apply visibility settings where supported.
            // Aspose.Cells Series does not expose an IsVisible property in all versions,
            // so this block is kept for future compatibility.
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                var series = chart.NSeries[i];
                if (visibilityPreferences.TryGetValue(i, out bool isVisible))
                {
                    // If the API provides a visibility property, set it here.
                    // Example: series.IsVisible = isVisible;
                }
            }

            // Save the workbook with the updated chart
            workbook.Save("WaterfallToggleVisibility.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
