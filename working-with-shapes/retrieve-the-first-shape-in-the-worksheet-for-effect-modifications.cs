// Title: Aspose.Cells .NET – Retrieve the First Shape and Apply a Shadow Effect
// Description: Creates a workbook, adds a rectangle shape, accesses the first shape via the Shapes collection (sheet.Shapes[0]), modifies its ShadowEffect (blur, distance, angle, transparency) and saves the file. Demonstrates quick shape retrieval and visual styling with Aspose.Cells for C#.
// Keywords: Aspose.Cells retrieve first shape | C# shape collection indexer | modify shape shadow effect | ShadowEffect blur distance angle | Aspose.Cells shape styling | add rectangle shape Aspose.Cells | .NET spreadsheet shape manipulation
// Common Searches: how to get the first shape in Aspose.Cells worksheet | set shadow blur and angle for a shape using Aspose.Cells C# | access shape collection and change visual effects in .NET | Aspose.Cells shape shadow properties example
// Developer Intent: The developer needs to locate the first shape in a worksheet and customize its shadow appearance programmatically.
// Use Cases: Add a rectangle to a worksheet, retrieve it with sheet.Shapes[0], and configure custom shadow blur, distance, angle, and transparency. | Check for the SolidFillColor property on the retrieved shape and apply a fill color before saving the workbook. | Apply a consistent shadow style to the first shape or to all shapes after retrieving them from the Shapes collection.
// AI Prompts: Write C# code that fetches the first shape from an Aspose.Cells worksheet and sets its ShadowEffect properties (Blur, Distance, Angle, Transparency). | Show how to detect if a shape supports SolidFillColor in Aspose.Cells and assign a color after retrieving the shape. | Provide a loop that iterates through every shape in a worksheet and applies identical shadow settings using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle shape, accesses the first shape via the Shapes collection (sheet.Shapes[0]), modifies its ShadowEffect (blur, distance, angle, transparency) and saves the file. Demonstrates quick shape retrieval and visual styling with Aspose.Cells for C#.
    class RetrieveFirstShapeDemo
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape rect = sheet.Shapes.AddRectangle(1, 0, 0, 0, 100, 100);

            // Retrieve the first shape in the worksheet using the ShapeCollection indexer
            Shape firstShape = sheet.Shapes[0];

            // Modify the shadow effect of the retrieved shape
            ShadowEffect shadow = firstShape.ShadowEffect;
            shadow.Blur = 5;               // Blur radius
            shadow.Distance = 10;          // Distance from the shape
            shadow.Angle = 45;             // Direction in degrees
            shadow.Transparency = 0.5;     // 50% transparent

            // Additional visual change: set solid fill color (if supported by the library version)
            // Uncomment the following line if SolidFillColor property is available in your Aspose.Cells version
            // firstShape.Fill.SolidFillColor = Color.LightBlue;

            // Save the workbook with the modified shape
            workbook.Save("FirstShapeEffectDemo.xlsx");
        }
    }
}
