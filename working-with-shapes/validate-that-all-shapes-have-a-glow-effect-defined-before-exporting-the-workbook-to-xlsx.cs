// Title: Validate & Apply Default Glow Effects to Shapes and Chart Series in Aspose.Cells (.NET)
// Description: C# sample that creates a workbook, adds rectangle and auto‑shapes, inserts a column chart, then checks every worksheet shape and each chart series for a glow effect. Missing glows are automatically set to a default color before the file is saved as XLSX.
// Keywords: Aspose.Cells glow effect | C# shape glow validation | default glow Aspose.Cells | chart series shape glow | apply glow to worksheet shapes | Aspose.Cells .NET example | export workbook with glow | shape property collection glow
// Common Searches: how to ensure all shapes have a glow in Aspose.Cells | set default glow for chart series using Aspose.Cells .NET | validate shape glow before saving workbook | add missing glow effect to Aspose.Cells shapes | C# Aspose.Cells shape glow example
// Developer Intent: Guarantee that every shape and chart series in a workbook has a defined glow effect before exporting to XLSX.
// Use Cases: Iterate through worksheet Shapes collection and assign a yellow glow when Glow.Size is zero or the Glow object is null. | Examine each Chart's Series ShapePropertyCollection and add an orange glow if no glow effect exists. | Save the workbook after all visual elements have a consistent glow, ensuring uniform appearance in the exported file.
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheet shapes and applies a default glow if missing. | Provide a method to check chart series shape properties and add a default glow effect before saving the workbook. | Create a reusable function that validates and sets glow effects for both generic shapes and chart series in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsGlowValidation
{
    // C# sample that creates a workbook, adds rectangle and auto‑shapes, inserts a column chart, then checks every worksheet shape and each chart series for a glow effect. Missing glows are automatically set to a default color before the file is saved as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape
            Shape rect = sheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);
            // No glow set initially

            // Add an auto shape
            Shape autoShape = sheet.Shapes.AddAutoShape(AutoShapeType.RoundedRectangle, 2, 0, 2, 0, 120, 80);
            // Set a glow effect for this shape (already defined)
            autoShape.Glow.Size = 12;
            autoShape.Glow.Color = workbook.CreateCellsColor();
            autoShape.Glow.Color.Color = Color.Green;

            // Add a sample chart (charts also contain shapes)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
            Chart chart = sheet.Charts[chartIdx];
            sheet.Cells["A1"].PutValue("Category 1");
            sheet.Cells["A2"].PutValue("Category 2");
            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            chart.NSeries.Add("B1:B2", true);
            chart.NSeries.CategoryData = "A1:A2";

            // Validate glow effect for all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // For generic shapes, check the Glow.Size (0 means not defined)
                if (shape.Glow == null || shape.Glow.Size == 0)
                {
                    // Define a default glow effect
                    shape.Glow.Size = 8; // default radius
                    shape.Glow.Color = workbook.CreateCellsColor();
                    shape.Glow.Color.Color = Color.Yellow;
                }
            }

            // Validate glow effect for chart series shapes
            foreach (Chart ch in sheet.Charts)
            {
                foreach (Series ser in ch.NSeries)
                {
                    ShapePropertyCollection spc = ser.ShapeProperties;
                    if (!spc.HasGlowEffect())
                    {
                        // Define a default glow effect for the series shape
                        spc.GlowEffect.Size = 6;
                        spc.GlowEffect.Color = workbook.CreateCellsColor();
                        spc.GlowEffect.Color.Color = Color.Orange;
                    }
                }
            }

            // Save the workbook to XLSX (lifecycle: save)
            workbook.Save("ValidatedGlowShapes.xlsx", SaveFormat.Xlsx);
        }
    }
}
