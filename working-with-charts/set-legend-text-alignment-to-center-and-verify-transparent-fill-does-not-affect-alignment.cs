// Title: Set chart legend to bottom center and apply 50% transparent fill using Aspose.Cells for .NET
// Common Searches: how to position a chart legend at the bottom center using Aspose.Cells in C# | apply 50% transparent fill to a chart legend without changing its alignment in Aspose.Cells | check that legend remains at bottom after setting fill transparency in an Aspose.Cells chart | Aspose.Cells example for setting legend position and transparent fill
// Tags: bottom legend alignment Aspose.Cells | legend area fill transparency C# | maintain legend position after fill change Aspose.Cells | column chart legend formatting Aspose.Cells | transparent legend fill without offset Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds data, inserts a column chart, positions the legend at the bottom (centered horizontally), sets the legend's area fill to 50% transparency, asserts that the legend stays at the bottom, and saves the file as LegendAlignmentTransparentFill.xlsx.
class LegendAlignmentExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------- Legend Alignment --------------------
            // Position the legend at the bottom (centered horizontally)
            chart.Legend.Position = LegendPositionType.Bottom;

            // -------------------- Transparent Fill Verification --------------------
            // Apply a transparent fill to the legend (50% transparency)
            // Legend does not expose FillFormat directly; use the Area object.
            chart.Legend.Area.FillFormat.Transparency = 0.5; // 0 = opaque, 1 = fully transparent

            // Verify that the legend's position remains Bottom after applying transparency
            Debug.Assert(chart.Legend.Position == LegendPositionType.Bottom,
                "Legend position should remain Bottom (centered) even with transparent fill.");

            // Save the workbook to a file
            string outputPath = "LegendAlignmentTransparentFill.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
