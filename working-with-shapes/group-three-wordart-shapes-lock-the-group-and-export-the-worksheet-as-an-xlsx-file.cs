// Title: Group and lock WordArt shapes then export the worksheet to XLSX with Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds three WordArt text effects, groups them into a single shape, locks the group, and saves the workbook as an XLSX file using Aspose.Cells. | Show a complete example of creating WordArt shapes, combining them into a locked group, and exporting the worksheet to an XLSX file with Aspose.Cells for .NET.
// Common Searches: aspacells c# combine WordArt objects and prevent editing | export Excel file with locked shape collection using Aspose.Cells | add text effect WordArt to worksheet and protect it in C# | C# Aspose.Cells example for grouping and locking shapes before saving
// Tags: WordArt shape grouping Aspose.Cells | locking shape groups in Excel C# | saving workbook as XLSX with Aspose.Cells | adding text effect shapes programmatically | protecting grouped shapes Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, inserts three WordArt text‑effect shapes, groups them, locks the group to prevent editing, and saves the worksheet as GroupedWordArt.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Add three WordArt shapes to the worksheet.
            // The AddTextEffect method requires width and height parameters after the cell range.
            var shape1 = worksheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset effect
                "Hello",                         // text
                "Arial",                         // font name
                36,                              // font size
                false,                           // bold
                false,                           // italic
                2, 2, 5, 5,                      // upper-left row/col, lower-right row/col
                200, 50);                        // width, height (in points)

            var shape2 = worksheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect2,
                "World",
                "Calibri",
                36,
                false,
                false,
                6, 2, 9, 5,
                200, 50);

            var shape3 = worksheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect3,
                "Aspose",
                "Times New Roman",
                36,
                false,
                false,
                10, 2, 13, 5,
                200, 50);

            // Group the three WordArt shapes
            Shape[] shapesToGroup = new Shape[] { shape1, shape2, shape3 };
            var groupShape = worksheet.Shapes.Group(shapesToGroup);

            // Lock the group to prevent editing
            groupShape.IsLocked = true;

            // Export the worksheet as an XLSX file
            workbook.Save("GroupedWordArt.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
