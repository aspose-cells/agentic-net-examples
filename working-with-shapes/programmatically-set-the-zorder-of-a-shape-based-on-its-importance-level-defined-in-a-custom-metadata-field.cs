// Title: Set worksheet shape Z‑order in Aspose.Cells C# using custom Importance metadata
// Description: Demonstrates how to add shapes to a worksheet, store an "Importance" value in each shape's Name property, parse that value, and call ToFrontOrBack to bring high‑importance shapes forward and send low‑importance shapes backward before saving the workbook.
// Keywords: Aspose.Cells shape Z-order | C# Aspose.Cells shape layering | custom metadata shape ordering | ToFrontOrBack Aspose.Cells | bring shape to front Aspose.Cells | send shape to back worksheet | programmatic shape priority
// Common Searches: Aspose.Cells change shape Z‑order by custom field | C# set shape layering based on importance in Excel | How to bring a shape to front in Aspose.Cells | Send low priority shape to back Aspose.Cells C# | Use shape Name property for metadata Aspose.Cells
// Developer Intent: Reorder worksheet shapes programmatically according to an importance level stored in a custom metadata field.
// Use Cases: Display critical annotations above charts and images in auto‑generated reports. | Prevent decorative graphics from covering data cells by sending them to the back. | Maintain a consistent visual hierarchy when multiple overlapping shapes are added via code.
// AI Prompts: Generate C# code that reads a numeric "Priority" custom property from each shape in an Aspose.Cells worksheet and reorders the shapes so higher priority shapes appear on top. | Show an example that extracts XML‑based custom metadata from shapes and uses it to set their Z‑order in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Demonstrates how to add shapes to a worksheet, store an "Importance" value in each shape's Name property, parse that value, and call ToFrontOrBack to bring high‑importance shapes forward and send low‑importance shapes backward before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add three shapes with a custom metadata field "Importance"
                // For demonstration we store the importance level in the shape's Name property
                Shape shapeA = sheet.Shapes.AddRectangle(5, 5, 100, 50, 0, 0);
                shapeA.Name = "Importance:Low";

                Shape shapeB = sheet.Shapes.AddRectangle(30, 30, 100, 50, 0, 0);
                shapeB.Name = "Importance:Medium";

                Shape shapeC = sheet.Shapes.AddRectangle(60, 60, 100, 50, 0, 0);
                shapeC.Name = "Importance:High";

                // Adjust Z‑order based on the importance level
                // Higher importance => bring to front, lower => send to back
                // Iterate backwards to avoid collection modification issues
                for (int i = sheet.Shapes.Count - 1; i >= 0; i--)
                {
                    Shape shp = sheet.Shapes[i];

                    // Extract the importance value from the Name (e.g., "Importance:High")
                    string[] parts = shp.Name.Split(':');
                    if (parts.Length != 2) continue; // skip if format is unexpected

                    string level = parts[1].Trim().ToLowerInvariant();

                    // Apply Z‑order change
                    if (level == "high")
                    {
                        // Bring to front
                        shp.ToFrontOrBack(1);
                    }
                    else if (level == "low")
                    {
                        // Send to back
                        shp.ToFrontOrBack(-1);
                    }
                    // "medium" – no change
                }

                // Save the workbook
                string outputPath = "ZOrderByImportance.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
