// Title: Create a WordArt TextEffect shape in an Excel worksheet with Aspose.Cells for .NET
// Common Searches: how to insert WordArt TextEffect into an Excel file using Aspose.Cells C# | Aspose.Cells add shape with preset text effect and custom font | C# generate Excel workbook containing WordArt shape programmatically | save Excel workbook with TextEffect shape using Aspose.Cells .NET | set position and size of TextEffect shape in Aspose.Cells worksheet
// Tags: Aspose.Cells add TextEffect shape | C# WordArt shape Excel worksheet | preset text effect WordArt Aspose.Cells | custom font size TextEffect shape | save workbook with shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a WordArt (TextEffect) shape with specified text, font, size, and positioning to the first worksheet, and saves the file as an XLSX workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt (TextEffect) shape to the worksheet
            Shape wordArt = worksheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset text effect
                "Aspose.Cells",                  // text
                "Arial",                         // font name
                36,                              // font size
                false,                           // bold
                false,                           // italic
                5,                               // left position (points)
                5,                               // top position (points)
                400,                             // width (points)
                100,                             // height (points)
                0,                               // rotation
                0                                // text direction
            );

            // Note: In some versions of Aspose.Cells the TextEffectFormat
            // and WordArt style APIs may not be available. The shape is
            // created successfully and can be saved without additional
            // formatting.

            // Save the workbook
            string outputFile = "WordArtShadowed.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
