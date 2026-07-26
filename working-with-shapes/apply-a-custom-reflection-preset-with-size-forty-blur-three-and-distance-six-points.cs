// Title: Apply Custom Reflection (size 40, blur 3, distance 6) to a Rectangle Shape with Aspose.Cells for .NET
// Description: Creates a new workbook, inserts a rectangle shape, and configures its ReflectionEffect to a custom preset (size = 40 %, blur = 3 pt, distance = 6 pt) before saving the file as CustomReflection.xlsx.
// Keywords: Aspose.Cells reflection effect | custom reflection preset C# | set shape blur distance | Rectangle shape visual style | .NET spreadsheet graphics | Aspose.Cells shape effects
// Common Searches: Aspose.Cells set custom reflection size blur distance | C# add rectangle with reflection effect in Excel | How to configure ReflectionEffect properties in Aspose.Cells | Apply custom reflection preset to shape using Aspose.Cells
// Developer Intent: Add a rectangle to a worksheet and modify its read‑only ReflectionEffect to a custom preset (size 40 %, blur 3 pt, distance 6 pt) then export the workbook.
// Use Cases: Design reports with highlighted sections that use subtle reflective styling for better visual hierarchy. | Generate marketing workbooks where product images are framed by reflective rectangles to draw attention. | Automate bulk workbook creation that requires consistent custom reflection settings across multiple shapes.
// AI Prompts: Show C# code to change the reflection size, blur, and distance of an existing shape to arbitrary values using Aspose.Cells. | Demonstrate how to apply the same custom reflection preset to every shape on a worksheet with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a rectangle shape, and configures its ReflectionEffect to a custom preset (size = 40 %, blur = 3 pt, distance = 6 pt) before saving the file as CustomReflection.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset Y, upper left offset X, width, height
            Shape rectangle = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 200, 100);

            // Access the existing ReflectionEffect (read‑only property) and configure it
            ReflectionEffect reflection = rectangle.Reflection;
            reflection.Type = ReflectionEffectType.Custom; // custom preset
            reflection.Size = 40;    // size in percentage
            reflection.Blur = 3;     // blur radius in points
            reflection.Distance = 6; // distance in points

            // Save the workbook with the applied reflection effect
            workbook.Save("CustomReflection.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
