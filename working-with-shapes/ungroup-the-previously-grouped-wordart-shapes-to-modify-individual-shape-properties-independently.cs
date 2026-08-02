// Title: Separate Grouped WordArt and Adjust TextEffect in Aspose.Cells (C#)
// Description: Shows how to add two WordArt objects to a worksheet, group them, then call GroupShape.Ungroup to break the group so each shape's TextEffect format—font name, size, bold, italic—can be modified independently before saving the workbook.
// Keywords: Aspose.Cells | C# WordArt ungroup | GroupShape.Ungroup | TextEffectFormat | modify WordArt font | Excel shape properties | Aspose.Cells .NET example | shape grouping | individual WordArt styling
// Common Searches: Aspose.Cells ungroup WordArt C# | Edit TextEffect of WordArt after grouping | GroupShape.Ungroup method example | Change font of individual WordArt in Excel using Aspose | C# code to separate grouped shapes Aspose.Cells
// Developer Intent: Break a grouped WordArt collection so each shape can be styled separately.
// Use Cases: Create a dashboard where headings are initially grouped for placement, then need distinct font styles. | Automate report generation that requires different WordArt effects after layout adjustments. | Programmatically adjust bold, italic, and font attributes of individual WordArt objects after they were grouped for alignment.
// AI Prompts: Generate C# Aspose.Cells code that groups two WordArt shapes, ungroups them, and sets one to bold Calibri 18pt and the other to italic Times New Roman 20pt. | Show an example of accessing TextEffectFormat of ungrouped WordArt shapes in Aspose.Cells and updating font properties. | Explain the steps and best practices for using GroupShape.Ungroup to modify individual WordArt shapes in an Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add two WordArt objects to a worksheet, group them, then call GroupShape.Ungroup to break the group so each shape's TextEffect format—font name, size, bold, italic—can be modified independently before saving the workbook.
class UngroupWordArtDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add two WordArt shapes with different preset styles
        Shape wordArt1 = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Hello",                         // text
            1, 1,                            // upper left row, column
            200, 50,                         // height, width
            0, 0);                           // image width, image height (not used for WordArt)

        Shape wordArt2 = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle2,
            "World",
            5, 1,
            200, 50,
            0, 0);

        // Group the two WordArt shapes into a single GroupShape
        GroupShape groupShape = worksheet.Shapes.Group(new Shape[] { wordArt1, wordArt2 });

        // Ungroup the shapes using the GroupShape.Ungroup method
        groupShape.Ungroup();

        // After ungrouping, modify each WordArt shape independently
        if (wordArt1.IsWordArt)
        {
            // Access the TextEffect format and change its properties
            TextEffectFormat effect1 = wordArt1.TextEffect;
            effect1.FontBold = true;
            effect1.FontName = "Calibri";
            effect1.FontSize = 18;
        }

        if (wordArt2.IsWordArt)
        {
            TextEffectFormat effect2 = wordArt2.TextEffect;
            effect2.FontItalic = true;
            effect2.FontName = "Times New Roman";
            effect2.FontSize = 20;
        }

        // Save the workbook with the modified shapes
        workbook.Save("UngroupedWordArtDemo.xlsx");
    }
}
