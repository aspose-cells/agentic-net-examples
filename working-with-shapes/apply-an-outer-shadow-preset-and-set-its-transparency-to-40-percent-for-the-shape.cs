// Title: Add an outer shadow with 40% transparency to a rectangle shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape on a worksheet and set its Shadow.Preset to OuterShadow with 40% transparency using Aspose.Cells C#. | Use reflection to detect the Shape.Shadow property, then enable the shadow, assign the OuterShadow preset, and set Transparency to 0.4. | Save the workbook as an .xlsx file after applying the outer shadow effect to the shape.
// Common Searches: how to apply outer shadow to a shape in Aspose.Cells C# | set shape shadow transparency to 40 percent using Aspose.Cells .NET | use reflection to access Shape.Shadow property in older Aspose.Cells versions | add rectangle shape with shadow effect in Excel file with Aspose.Cells | Aspose.Cells C# example for configuring shape shadow properties
// Tags: Aspose.Cells shape outer shadow | set shape shadow transparency C# | reflection access Shape.Shadow Aspose.Cells | add rectangle shape Excel Aspose.Cells | save workbook with shaped shadow .xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, inserts a rectangle shape, uses reflection to safely access the Shape.Shadow property, sets the shadow preset to OuterShadow, applies 40% transparency, makes the shadow visible, and saves the file as ShapeWithOuterShadow.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper‑left row, upper‑left column,
                // upper‑left row offset (points), upper‑left column offset (points),
                // width (points), height (points)
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    2, // row index (zero‑based)
                    2, // column index (zero‑based)
                    0, // upper‑left row offset
                    0, // upper‑left column offset
                    120, // width
                    60   // height
                );

                // Apply an outer shadow preset to the shape (if supported by the library version)
                // The Shadow property is available in recent Aspose.Cells versions.
                // Guard against missing members using reflection to keep compatibility with older versions.
                var shadowProperty = typeof(Shape).GetProperty("Shadow");
                if (shadowProperty != null)
                {
                    var shadow = shadowProperty.GetValue(shape);
                    var presetProp = shadow.GetType().GetProperty("Preset");
                    var transparencyProp = shadow.GetType().GetProperty("Transparency");
                    var visibleProp = shadow.GetType().GetProperty("Visible");

                    if (presetProp != null && transparencyProp != null && visibleProp != null)
                    {
                        // Set shadow preset
                        var presetEnum = Enum.Parse(presetProp.PropertyType, "OuterShadow");
                        presetProp.SetValue(shadow, presetEnum);

                        // Set transparency (40%)
                        transparencyProp.SetValue(shadow, 0.40);

                        // Make shadow visible
                        visibleProp.SetValue(shadow, true);
                    }
                }

                // Define output file path
                string outputPath = "ShapeWithOuterShadow.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
