using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRichTextReplaceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access cell A1 and set initial rich text
            Cell cell = worksheet.Cells["A1"];
            cell.Value = "Hello World";

            // Apply formatting to the first part "Hello"
            FontSetting helloSetting = cell.Characters(0, 5);
            helloSetting.Font.IsBold = true;
            helloSetting.Font.Color = Color.Red;

            // Apply formatting to the second part "World"
            FontSetting worldSetting = cell.Characters(6, 5);
            worldSetting.Font.IsItalic = true;
            worldSetting.Font.Color = Color.Blue;

            // ----- Replace "World" with "Aspose" using SetCharacters -----
            // New text for the cell
            string newText = "Hello Aspose";
            cell.Value = newText;

            // Create new FontSetting array reflecting the new text portions
            FontSetting[] newSettings = new FontSetting[2];

            // Keep the original formatting for "Hello"
            newSettings[0] = new FontSetting(0, 5, workbook.Worksheets);
            newSettings[0].Font.IsBold = true;
            newSettings[0].Font.Color = Color.Red;

            // Apply formatting for the new word "Aspose"
            newSettings[1] = new FontSetting(6, 6, workbook.Worksheets); // "Aspose" length is 6
            newSettings[1].Font.IsItalic = true;
            newSettings[1].Font.Color = Color.Blue;

            // Set the new character formatting to the cell
            cell.SetCharacters(newSettings);

            // Save the workbook
            workbook.Save("RichTextReplaceDemo.xlsx");
        }
    }
}