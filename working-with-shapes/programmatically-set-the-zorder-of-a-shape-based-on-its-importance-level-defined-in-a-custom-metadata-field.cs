// Title: Set Shape Z‑Order by Custom Importance Metadata (AlternativeText) in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to read an "Importance" value from the AlternativeText property of worksheet shapes, sort the shapes by that value, and adjust their Z‑order with Shape.ToFrontOrBack so that higher‑importance shapes appear on top. The workbook is saved with the new stacking order.
// Keywords: Aspose.Cells shape Z-order | C# Aspose.Cells shape ordering | AlternativeText custom metadata | shape importance priority | ToFrontOrBack method | reorder worksheet shapes programmatically | Excel drawing objects Z‑order | set shape stacking order | custom metadata for shapes | Aspose.Cells .NET example
// Common Searches: Aspose.Cells change shape Z order based on metadata | C# move Excel shapes to front using AlternativeText | How to sort worksheet shapes by priority in Aspose.Cells | Set shape stacking order programmatically Aspose.Cells | ToFrontOrBack usage example C#
// Developer Intent: Reorder worksheet shapes according to an importance level stored in each shape’s AlternativeText using Aspose.Cells for .NET.
// Use Cases: Extract a numeric importance value from the AlternativeText of each shape and sort the shape collection before reordering. | Apply Shape.ToFrontOrBack with the extracted importance to move higher‑priority shapes forward in the Z‑order stack. | Save the workbook after reordering so the visual layout reflects the defined priorities. | Integrate dynamic shape layering in reports where business rules dictate visual prominence.
// AI Prompts: Generate C# code that reads an "Importance" field from Shape.AlternativeText and adjusts the Z‑order with ToFrontOrBack in Aspose.Cells. | Write a method to sort worksheet shapes by a numeric priority stored in AlternativeText and bring them to the front in order. | Explain how Shape.ToFrontOrBack works for moving shapes forward by a specific number of positions in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Demonstrates how to read an "Importance" value from the AlternativeText property of worksheet shapes, sort the shapes by that value, and adjust their Z‑order with Shape.ToFrontOrBack so that higher‑importance shapes appear on top. The workbook is saved with the new stacking order.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add three rectangle shapes with a custom metadata field stored in AlternativeText
                // The metadata represents an importance level (higher value = more important)
                Shape shapeA = sheet.Shapes.AddRectangle(5, 5, 100, 50, 0, 0);
                shapeA.AlternativeText = "Importance=2";

                Shape shapeB = sheet.Shapes.AddRectangle(30, 30, 100, 50, 0, 0);
                shapeB.AlternativeText = "Importance=5";

                Shape shapeC = sheet.Shapes.AddRectangle(60, 60, 100, 50, 0, 0);
                shapeC.AlternativeText = "Importance=1";

                // Copy shapes to a separate list to avoid modifying the collection during enumeration
                List<Shape> shapes = sheet.Shapes.Cast<Shape>().ToList();

                // Sort shapes by importance (higher importance later so they end up in front)
                shapes.Sort((s1, s2) =>
                {
                    int imp1 = GetImportance(s1);
                    int imp2 = GetImportance(s2);
                    return imp1.CompareTo(imp2);
                });

                // Bring shapes to front in order of increasing importance
                foreach (Shape shp in shapes)
                {
                    int importance = GetImportance(shp);
                    if (importance != 0)
                    {
                        // Move the shape forward by its importance value
                        shp.ToFrontOrBack(importance);
                    }
                }

                // Save the workbook with the updated Z‑order
                workbook.Save("ZOrderByImportance.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to extract importance from AlternativeText
        private static int GetImportance(Shape shape)
        {
            int importance = 0;
            string meta = shape.AlternativeText;
            if (!string.IsNullOrEmpty(meta) && meta.StartsWith("Importance=", StringComparison.OrdinalIgnoreCase))
            {
                int.TryParse(meta.Substring("Importance=".Length), out importance);
            }
            return importance;
        }
    }
}
