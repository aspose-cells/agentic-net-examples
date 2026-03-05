using System;
using Aspose.Cells;
using System.Drawing;

class RichTextCellDemo
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook wb = new Workbook();

        // Access the first worksheet
        Worksheet ws = wb.Worksheets[0];

        // Get the target cell
        Cell cell = ws.Cells["A1"];

        // Set initial plain text
        cell.PutValue("Hello World!");

        // Prepare font settings for different parts of the text
        FontSetting[] fontSettings = new FontSetting[2];

        // Format "Hello" (characters 0-4) as bold red
        fontSettings[0] = cell.Characters(0, 5);
        fontSettings[0].Font.IsBold = true;
        fontSettings[0].Font.Color = Color.Red;

        // Format "World!" (characters 6-11) as italic blue
        fontSettings[1] = cell.Characters(6, 6);
        fontSettings[1].Font.IsItalic = true;
        fontSettings[1].Font.Color = Color.Blue;

        // Apply the rich text formatting to the cell
        cell.SetCharacters(fontSettings);

        // Insert additional text at index 5 (after "Hello") while preserving existing formatting
        cell.InsertText(5, " Beautiful");

        // Verify that the cell now contains rich text
        bool isRich = cell.IsRichText();
        Console.WriteLine("Cell contains rich text: " + isRich);

        // Optionally retrieve the rich value object
        CellRichValue richValue = cell.GetRichValue();

        // Save the workbook (save rule)
        wb.Save("RichTextDemo.xlsx");
    }
}