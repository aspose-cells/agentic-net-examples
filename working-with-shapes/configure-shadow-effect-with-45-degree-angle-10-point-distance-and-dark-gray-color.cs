// Title: Add a 45‑degree dark gray shadow with 10‑point offset to a rectangle shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape shadow angle to 45 degrees | how to apply a dark gray shadow with specific distance to a shape in Aspose.Cells | configure shadow offset and color for rectangle shape in Aspose.Cells .NET | set shadow distance 10 points for shape using Aspose.Cells API | apply custom shadow effect to shapes in an Excel workbook with Aspose.Cells
// Tags: shadow angle setting Aspose.Cells | shadow distance setting Aspose.Cells | custom shadow color Aspose.Cells | rectangle shape shadow Aspose.Cells | shape shadow configuration Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, adding a rectangle shape, and applying a shadow effect with a 45° angle, 10‑point distance, and dark gray color using Aspose.Cells for .NET, then saving the file.
class Program
{
    static void Main()
    {
        // Initialize a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to demonstrate the shadow effect
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 200, 100);

        // Configure the shadow effect
        ShadowEffect shadow = shape.ShadowEffect;
        shadow.Angle = 45;          // 45 degree angle
        shadow.Distance = 10;       // 10 points distance

        // Create a dark gray CellsColor and assign it to the shadow
        CellsColor darkGray = workbook.CreateCellsColor();
        darkGray.Color = Color.DarkGray; // System.Drawing.Color
        shadow.Color = darkGray;

        // Optional: adjust other shadow properties if desired
        // shadow.Blur = 5;
        // shadow.Transparency = 0.2;

        // Save the workbook
        string outputPath = "ShadowEffectDemo.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine($"Workbook saved to '{outputPath}'. Shadow effect applied: Angle=45°, Distance=10pt, Color=DarkGray.");
    }
}
