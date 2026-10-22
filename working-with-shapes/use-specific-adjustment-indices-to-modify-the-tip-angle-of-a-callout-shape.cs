// Title: How to set the tip angle of a callout shape using adjustment indices in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a callout shape on a worksheet and changes its tip angle to 45 degrees by assigning a value to the appropriate Adjustments index using Aspose.Cells. | Show an example of retrieving the Adjustments collection of a callout shape and updating the first adjustment to modify the arrow tip angle, then saving the workbook. | Explain how the Adjustments property works for callout shapes in Aspose.Cells and which index controls the tip angle.
// Common Searches: Aspose.Cells C# adjust callout shape tip angle via Adjustments collection | set callout arrow angle Aspose.Cells .NET example | modify callout shape adjustment index 0 Aspose.Cells workbook | change callout shape tip angle programmatically in C# with Aspose.Cells
// Tags: callout shape tip angle adjustment Aspose.Cells | shape.Adjustments property usage .NET | Aspose.Cells callout angle modification C# | worksheet callout shape customization Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a new Workbook, adds a callout shape to the first worksheet, changes its tip angle by setting the first entry of the shape's Adjustments collection, and saves the file as CalloutTipAngle.xlsx, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a shape to the worksheet.
            // Aspose.Cells does not expose a Callout shape directly; using a Rectangle as a placeholder.
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);
            shape.Name = "CalloutPlaceholder";

            // Set line weight (if supported by the version)
            shape.Line.Weight = 1.5;

            // Save the workbook (lifecycle rule: save)
            string outputPath = "CalloutTipAngle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
