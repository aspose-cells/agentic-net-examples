// Title: Create a WordArt (TextEffect) shape in an Excel worksheet and save the workbook as XLSX with Aspose.Cells for .NET
// Common Searches: how to insert a WordArt TextEffect shape into an Excel file using Aspose.Cells C# | Aspose.Cells example for adding WordArt to a worksheet | saving an Excel workbook that contains a WordArt shape with Aspose.Cells .NET | setting font and size for a WordArt shape in Aspose.Cells C# code | why adjustment values for WordArt curvature are not available in Aspose.Cells
// Tags: add TextEffect shape Aspose.Cells C# | WordArt shape Excel Aspose.Cells | save workbook with shapes XLSX Aspose.Cells | Aspose.Cells shape API limitation adjustment values | Excel WordArt curvature unsupported Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates creating a new Workbook, adding a WordArt (TextEffect) shape with specific text, font, size, and positioning to the first worksheet using Aspose.Cells for .NET, notes that curvature adjustment values are not supported, and saves the file as 'WordArtExample.xlsx' in XLSX format.
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

            // Add a WordArt shape (TextEffect) to the worksheet
            // Parameters: preset, text, font name, font size, bold, italic,
            // upper left row, upper left column, lower right row, lower right column,
            // width, height
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Aspose.Cells",
                "Arial",
                36,
                false,
                false,
                2, 1, 10, 5,
                200, 100);

            // Note: AdjustmentValues are not available in Aspose.Cells Shape API.
            // The default curvature of the WordArt will be used.

            // Save the workbook
            workbook.Save("WordArtExample.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
