// Title: C# Aspose.Cells function to check if a worksheet shape is a locked WordArt watermark
// Common Searches: how to determine if a shape is a locked WordArt watermark using Aspose.Cells C# | Aspose.Cells C# check shape.IsWordArt and shape.IsLocked properties | C# code to find locked WordArt watermarks in an Excel worksheet with Aspose.Cells | detect locked WordArt watermark shapes in Excel file using Aspose.Cells library | sample method to identify locked WordArt watermarks in an Aspose.Cells workbook
// Tags: detect locked WordArt shape Aspose.Cells | shape.IsWordArt usage C# | shape.IsLocked usage Aspose.Cells | scan worksheet shapes Aspose.Cells | custom shape helper C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example defines a static ShapeHelper.IsLockedWordArtWatermark method that returns true when a Shape's IsWordArt and IsLocked properties are both true, and demonstrates iterating through a worksheet's Shapes collection to locate such locked WordArt watermarks in an Excel file using Aspose.Cells for C#.
public static class ShapeHelper
{
    // Returns true if the supplied shape is a locked WordArt watermark.
    public static bool IsLockedWordArtWatermark(Shape shape)
    {
        // Guard against null references.
        if (shape == null)
            return false;

        // Aspose.Cells provides IsWordArt and IsLocked properties.
        // A shape qualifies as a locked WordArt watermark when both are true.
        return shape.IsWordArt && shape.IsLocked;
    }
}

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Ensure the input file exists; otherwise create a new workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
                workbook.Save(inputPath);   // optionally save the placeholder
            }

            // Work with the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Scan all shapes on the sheet.
            foreach (Shape shape in sheet.Shapes)
            {
                if (ShapeHelper.IsLockedWordArtWatermark(shape))
                {
                    // Example action: output the shape name.
                    Console.WriteLine($"Locked WordArt watermark found: {shape.Name}");
                }
            }

            // Save the workbook if any modifications were made.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Release resources if needed.
            workbook?.Dispose();
        }
    }
}
