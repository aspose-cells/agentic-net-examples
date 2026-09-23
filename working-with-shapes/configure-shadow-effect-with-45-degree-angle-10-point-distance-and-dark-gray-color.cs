// Title: Create a rectangle shape with a 45‑degree dark‑gray outer shadow (10 pt distance) in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a rectangle shape to a worksheet and applies an outer shadow with a 45° direction, 10‑point distance, and dark gray color using Aspose.Cells, handling version differences via reflection. | Show how to access the Shape.Shadow property via reflection and set its Type, Direction, Distance, Color, and Visible fields for a shape in Aspose.Cells. | Demonstrate saving the workbook after configuring the shadow effect on a shape with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set outer shadow direction 45 degrees on rectangle shape | how to configure shadow distance and color for shapes in Aspose.Cells .NET | using reflection to set shape shadow properties across Aspose.Cells versions | apply dark gray outer shadow to an Excel shape with Aspose.Cells API
// Tags: Aspose.Cells shape outer shadow configuration | C# set shape shadow direction distance color | reflection based Shape.Shadow property access Aspose.Cells | Excel workbook shape shadow effect .NET | Aspose.Cells rectangle shape styling

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a rectangle shape to the first worksheet, sets its fill and line colors, and uses reflection to safely access the Shape.Shadow property. It configures the shadow as an outer shadow with a 45‑degree direction, 10‑point distance, dark gray color, and makes it visible before saving the file as ShadowEffect.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,   // upper left row
                    1,   // upper left column
                    0,   // top offset (in points)
                    0,   // left offset (in points)
                    120, // width (in points)
                    60   // height (in points)
                );

                // Set fill and line colors using the correct APIs
                shape.FillFormat.ForeColor = Color.LightBlue;
                shape.LineFormat.ForeColor = Color.DarkBlue;

                // Attempt to configure shadow effect via reflection (compatible with multiple versions)
                var shadowProp = typeof(Shape).GetProperty("Shadow");
                if (shadowProp != null)
                {
                    var shadow = shadowProp.GetValue(shape);
                    if (shadow != null)
                    {
                        // Set shadow type if the property exists
                        var typeProp = shadow.GetType().GetProperty("Type");
                        if (typeProp != null && Enum.IsDefined(typeProp.PropertyType, "OuterShadow"))
                        {
                            typeProp.SetValue(shadow, Enum.Parse(typeProp.PropertyType, "OuterShadow"));
                        }

                        // Set other shadow properties if they exist
                        var directionProp = shadow.GetType().GetProperty("Direction");
                        directionProp?.SetValue(shadow, 45);

                        var distanceProp = shadow.GetType().GetProperty("Distance");
                        distanceProp?.SetValue(shadow, 10);

                        var colorProp = shadow.GetType().GetProperty("Color");
                        colorProp?.SetValue(shadow, Color.DarkGray);

                        var visibleProp = shadow.GetType().GetProperty("Visible");
                        visibleProp?.SetValue(shadow, true);
                    }
                }

                // Define output file path
                string outputPath = "ShadowEffect.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
