// Title: Add WordArt with a built‑in style using the ShapeCollection.AddWordArt style‑name overload in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add WordArt using style name overload | How to apply a built‑in WordArt style when creating Excel with Aspose.Cells | ShapeCollection AddWordArt with predefined style parameter example | C# code to insert WordArt with style name into an Excel worksheet using Aspose.Cells
// Tags: ShapeCollection.AddWordArt style name overload | Aspose.Cells WordArt built‑in style | C# add WordArt to Excel worksheet | Excel WordArt formatting programmatically | apply predefined WordArt style Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first Worksheet's ShapeCollection, and adds a WordArt shape by calling the AddWordArt overload that accepts a built‑in style name. The style is applied directly during shape creation, after which the workbook is saved as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the shape collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add WordArt using the 8‑parameter overload
            Shape wordArt = shapes.AddWordArt(
                0,                     // index of the shape
                "Aspose.Cells",        // text to display
                2,                     // upper left row
                2,                     // upper left column
                0,                     // top offset (pixels)
                0,                     // left offset (pixels)
                400,                   // width (pixels)
                100);                  // height (pixels)

            // Apply basic formatting (optional)
            wordArt.TextEffect.FontName = "Arial";
            wordArt.TextEffect.FontSize = 36;

            // Save the workbook
            workbook.Save("WordArtStyleExample.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
