using System;
using Aspose.Cells;

class ChangeRichTextPortionFontSize
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put rich text into cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Hello World");

        // Select the portion "World" (starts at index 6, length 5)
        FontSetting richTextPortion = cell.Characters(6, 5);

        // Change its font size to 12 points
        richTextPortion.Font.Size = 12;

        // Save the workbook
        workbook.Save("RichTextPortionFontSize.xlsx");
    }
}