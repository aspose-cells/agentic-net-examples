using System;
using System.Drawing;
using Aspose.Cells;

class RichTextExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get cell A1
        Cell cell = sheet.Cells["A1"];

        // Set initial text in the cell
        cell.PutValue("Hello World");

        // Apply rich formatting:
        // "Hello" (characters 0-4) -> bold and red
        FontSetting fsHello = cell.Characters(0, 5);
        fsHello.Font.IsBold = true;
        fsHello.Font.Color = Color.Red;

        // "World" (characters 6-10) -> italic and blue
        FontSetting fsWorld = cell.Characters(6, 5);
        fsWorld.Font.IsItalic = true;
        fsWorld.Font.Color = Color.Blue;

        // Apply the character settings to the cell
        cell.SetCharacters(new FontSetting[] { fsHello, fsWorld });

        // Change the color of the "World" segment to green
        FontSetting[] allSegments = cell.GetCharacters();
        foreach (FontSetting segment in allSegments)
        {
            if (segment.StartIndex == 6 && segment.Length == 5)
            {
                segment.Font.Color = Color.Green;
            }
        }

        // Insert additional text "Beautiful " before "World"
        // Index 6 is the position right after "Hello "
        cell.InsertText(6, "Beautiful ");

        // Save the workbook to an XLSX file
        workbook.Save("RichTextDemo.xlsx");
    }
}