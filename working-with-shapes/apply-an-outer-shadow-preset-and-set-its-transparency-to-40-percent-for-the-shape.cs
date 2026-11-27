// Title: How to apply an outer shadow preset with 40% transparency to a rectangle shape using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set outer shadow preset on worksheet shape | apply 40 percent transparency to shape shadow in Aspose.Cells .NET | Aspose.Cells example for adding rectangle shape with offset bottom shadow | how to configure shadow effect transparency for Excel shape using Aspose.Cells | set shadow preset type and transparency for shape in C# workbook
// Tags: apply outer shadow preset using Aspose.Cells | configure shape shadow transparency in .NET | offset bottom shadow type for Excel shape | add rectangle shape with shadow effect C# | Aspose.Cells shadow effect properties example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, inserts a rectangle shape on the first worksheet, accesses its ShadowEffect, sets the preset type to OffsetBottom (an outer shadow), applies 40% transparency (0.4), and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        // Path to the output Excel file
        const string outputPath = "ShadowOuterPreset.xlsx";

        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftPixelRow, upperLeftPixelColumn, width, height
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 200, 100);

            // Access the shadow effect of the shape
            ShadowEffect shadow = shape.ShadowEffect;

            // Apply an outer shadow preset (using an offset type as outer shadow)
            shadow.PresetType = PresetShadowType.OffsetBottom;

            // Set the transparency to 40% (0.4 in the 0.0‑1.0 range)
            shadow.Transparency = 0.4;

            // Optional: adjust other shadow properties for better visibility
            // shadow.Color = workbook.CreateCellsColor(); // default black
            // shadow.Color.Color = System.Drawing.Color.Gray;
            // shadow.Blur = 20;
            // shadow.Distance = 10;

            // Save the workbook
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            Console.WriteLine("The shape now has an outer shadow preset with 40% transparency.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while applying the shadow effect:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
