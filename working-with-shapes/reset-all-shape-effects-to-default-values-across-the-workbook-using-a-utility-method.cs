// Title: Reset all chart series shape effects (3D, glow, shadow) in an Aspose.Cells workbook using C#
// Description: The sample builds a workbook, inserts data, a column chart and a few drawing shapes, applies custom 3D, glow and shadow effects to a chart series, then runs a utility method that iterates over every worksheet, chart and series to clear those effects with ClearFormat3D, ClearGlowEffect and ClearShadowEffect before saving the file.
// Keywords: Aspose.Cells | C# | .NET | reset chart shape effects | clear 3D format Aspose.Cells | remove glow effect chart series | remove shadow effect chart series | ShapePropertyCollection ClearFormat3D | ShapePropertyCollection ClearGlowEffect | ShapePropertyCollection ClearShadowEffect | Excel automation | workbook cleanup
// Common Searches: How to clear 3D, glow and shadow effects from all chart series in Aspose.Cells C# | Aspose.Cells utility to reset shape effects for every chart in a workbook | Remove custom visual effects from Excel charts using Aspose.Cells .NET | Clear chart series formatting programmatically with Aspose.Cells | Reset shape properties across all worksheets in an Aspose.Cells workbook
// Developer Intent: Remove all custom 3D, glow, and shadow formatting from chart series throughout a workbook.
// Use Cases: Standardize chart appearance before exporting to PDF or sharing with clients. | Strip third‑party visual styles to comply with corporate branding guidelines. | Clean up charts after copying between workbooks to avoid rendering anomalies. | Prepare a workbook for automated testing by ensuring default visual settings.
// AI Prompts: Generate C# code that loops through every worksheet, chart, and series in an Aspose.Cells Workbook and calls ClearFormat3D, ClearGlowEffect, and ClearShadowEffect on each series. | Show how to create a reusable method in Aspose.Cells for .NET that resets shape effects on chart series and optionally on drawing shapes. | Explain the differences between ClearFormat3D, ClearGlowEffect, and ClearShadowEffect and when each should be used in Excel automation.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeEffectReset
{
    // The sample builds a workbook, inserts data, a column chart and a few drawing shapes, applies custom 3D, glow and shadow effects to a chart series, then runs a utility method that iterates over every worksheet, chart and series to clear those effects with ClearFormat3D, ClearGlowEffect and ClearShadowEffect before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Populate workbook with sample data, charts and shapes
            PopulateSampleContent(workbook);

            // Reset all shape effects to their default values across the workbook
            ResetAllShapeEffects(workbook);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("Workbook_With_ResetShapeEffects.xlsx", SaveFormat.Xlsx);
        }

        // Adds sample data, a chart and a few shapes to demonstrate the reset utility
        private static void PopulateSampleContent(Workbook workbook)
        {
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the chart
            sheet.Cells["A1"].PutValue("Category 1");
            sheet.Cells["A2"].PutValue("Category 2");
            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B1:B2", true);
            chart.NSeries.CategoryData = "A1:A2";

            // Apply some effects to the first series (so we have something to clear later)
            Series series = chart.NSeries[0];
            ShapePropertyCollection spc = series.ShapeProperties;
            spc.Format3D.SurfaceMaterialType = PresetMaterialType.WarmMatte;
            spc.ShadowEffect.Size = 1.5;
            spc.GlowEffect.Size = 30;

            // Add a few drawing shapes
            sheet.Shapes.AddRectangle(2, 0, 2, 0, 50, 50);
            sheet.Shapes.AddOval(8, 0, 2, 0, 60, 60);
        }

        // Utility method that resets all shape effects (3D, glow, shadow) for every chart series
        // in every worksheet of the provided workbook.
        private static void ResetAllShapeEffects(Workbook workbook)
        {
            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in ws.Charts)
                {
                    // Iterate through all series of the chart
                    foreach (Series series in chart.NSeries)
                    {
                        // Access the shape property collection of the series
                        ShapePropertyCollection shapeProps = series.ShapeProperties;

                        // Clear 3D format, glow effect and shadow effect (rule: use Clear* methods)
                        shapeProps.ClearFormat3D();
                        shapeProps.ClearGlowEffect();
                        shapeProps.ClearShadowEffect();
                    }
                }
            }
        }
    }
}
