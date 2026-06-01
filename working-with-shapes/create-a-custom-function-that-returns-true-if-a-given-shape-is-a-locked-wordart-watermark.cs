using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

public static class ShapeHelper
{
    // Returns true if the shape is a WordArt and is locked (commonly used for watermarks)
    public static bool IsLockedWordArtWatermark(Shape shape)
    {
        if (shape == null)
            return false;

        return shape.IsWordArt && shape.IsLocked;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Example: load a workbook only if the file exists
            string workbookPath = "sample.xlsx";

            if (File.Exists(workbookPath))
            {
                var workbook = new Workbook(workbookPath);
                var worksheet = workbook.Worksheets[0];

                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    bool isLockedWordArt = ShapeHelper.IsLockedWordArtWatermark(shape);
                    Console.WriteLine($"Shape \"{shape.Name}\": Locked WordArt = {isLockedWordArt}");
                }
            }
            else
            {
                Console.WriteLine($"Workbook file \"{workbookPath}\" not found. Skipping processing.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}