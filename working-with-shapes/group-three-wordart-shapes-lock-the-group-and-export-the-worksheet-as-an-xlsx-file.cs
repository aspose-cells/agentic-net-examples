// Title: Group and lock WordArt shapes then save the worksheet as XLSX using Aspose.Cells for .NET
// Common Searches: how to group WordArt shapes in a worksheet with Aspose.Cells C# | locking a shape group in Aspose.Cells before saving to XLSX | create WordArt text effects and export workbook as XLSX using Aspose.Cells | Aspose.Cells C# example for grouping and locking shapes | save Excel file with locked shape group using Aspose.Cells
// Tags: WordArt shape grouping Aspose.Cells | shape group lock Aspose.Cells | AddTextEffect text effect creation C# | export workbook with grouped shapes to XLSX | Aspose.Cells shape grouping API

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates three WordArt‑style text effect shapes, groups them into a single GroupShape, locks the group to prevent editing, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add three WordArt‑like shapes using AddTextEffect (WordArt is added as a text effect)
            Shape wordArt1 = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset text effect
                "First",                         // text
                "Arial",                         // font name
                36, false, false,                // font size, bold, italic
                2, 2, 0, 0, 200, 50);             // row, column, top, left, width, height

            Shape wordArt2 = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Second",
                "Arial",
                36, false, false,
                5, 2, 0, 0, 200, 50);

            Shape wordArt3 = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Third",
                "Arial",
                36, false, false,
                8, 2, 0, 0, 200, 50);

            // Group the three shapes
            Shape[] shapesToGroup = new Shape[] { wordArt1, wordArt2, wordArt3 };
            GroupShape group = sheet.Shapes.Group(shapesToGroup);

            // Lock the group to prevent editing (use IsLocked property for GroupShape)
            group.IsLocked = true;

            // Save the workbook
            string outputPath = "GroupedWordArt.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
