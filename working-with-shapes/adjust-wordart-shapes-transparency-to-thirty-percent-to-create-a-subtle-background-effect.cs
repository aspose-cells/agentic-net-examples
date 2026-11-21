// Title: Set 30% fill transparency on a WordArt (TextEffect) shape in an Excel worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set WordArt shape fill transparency to 30 percent | how to make TextEffect shape semi‑transparent in an Excel file using .NET | adjust opacity of WordArt background in Aspose.Cells workbook | C# code example for setting shape fill transparency in Excel with Aspose | make WordArt outline invisible Aspose.Cells .NET example
// Tags: shape fill transparency Aspose.Cells | WordArt TextEffect opacity .NET | Excel shape transparency C# | transparent WordArt background Aspose | hide shape outline Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a TextEffect (WordArt) shape, sets its fill transparency to 30 % (0.3) and makes the outline fully transparent, then saves the file as WordArtTransparency.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape (TextEffect) to the worksheet
            // Parameters: preset effect, text, font name, font size, isBold, isItalic,
            // left, top, width, height, shape rotation, shape depth (in points)
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Subtle Background",
                "Arial",
                48,
                false,
                false,
                0,
                0,
                400,
                100,
                0,   // rotation
                0);  // depth

            // Set the fill transparency to 30% (0.3). This makes the WordArt appear as a subtle background.
            wordArt.Fill.Transparency = 0.3;

            // Hide the outline by making it fully transparent
            wordArt.Line.Transparency = 1.0;
            // Alternatively, you could set the line color to transparent:
            // wordArt.Line.Color = Color.Transparent;

            // Save the workbook to a file
            string outputPath = "WordArtTransparency.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
