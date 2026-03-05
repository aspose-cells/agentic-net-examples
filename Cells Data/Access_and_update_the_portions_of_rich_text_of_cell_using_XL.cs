using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRichTextDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Set initial plain text in the cell
            cell.PutValue("Hello Aspose Cells");

            // Apply rich formatting to two parts of the text:
            // "Hello" -> bold red, "Aspose" -> italic blue
            FontSetting[] initialSettings = new FontSetting[2];
            initialSettings[0] = cell.Characters(0, 5);   // "Hello"
            initialSettings[0].Font.IsBold = true;
            initialSettings[0].Font.Color = Color.Red;

            initialSettings[1] = cell.Characters(6, 6);   // "Aspose"
            initialSettings[1].Font.IsItalic = true;
            initialSettings[1].Font.Color = Color.Blue;

            // Apply the initial rich formatting
            cell.SetCharacters(initialSettings);

            // Insert additional text " .NET" after the word "Cells"
            string currentText = cell.StringValue;
            int insertPos = currentText.IndexOf("Cells") + "Cells".Length;
            cell.InsertText(insertPos, " .NET");

            // Modify formatting of the newly inserted segment to green and underlined
            // The inserted segment starts at insertPos and has length 5 (" .NET")
            cell.Characters(insertPos, 5).Font.Color = Color.Green;
            cell.Characters(insertPos, 5).Font.Underline = FontUnderlineType.Single;

            // Re-apply all character settings to preserve existing formatting
            cell.SetCharacters(cell.GetCharacters());

            // Save the workbook in XLSX format
            workbook.Save("RichTextUpdated.xlsx");
        }
    }
}