using System;
using System.Drawing;
using Aspose.Cells;

class SerializeRichTextPortion
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell A1
        Cell cell = worksheet.Cells["A1"];

        // Set plain text value
        cell.PutValue("Hello World");

        // Apply rich text formatting:
        // "Hello" (first 5 characters) -> bold and red
        FontSetting helloPortion = cell.Characters(0, 5);
        helloPortion.Font.IsBold = true;
        helloPortion.Font.Color = Color.Red;

        // "World" (characters starting at index 6, length 5) -> italic and blue
        FontSetting worldPortion = cell.Characters(6, 5);
        worldPortion.Font.IsItalic = true;
        worldPortion.Font.Color = Color.Blue;

        // Verify that the cell contains rich text
        bool isRich = cell.IsRichText();
        Console.WriteLine("IsRichText: " + isRich);

        // Serialize the cell (including its rich text portions) to JSON
        string cellJson = cell.ToJson();

        // Output the JSON string – it can be stored and reused later
        Console.WriteLine("Cell JSON with rich text portions:");
        Console.WriteLine(cellJson);
    }
}