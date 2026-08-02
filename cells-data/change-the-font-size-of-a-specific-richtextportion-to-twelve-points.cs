using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a cell value that contains text
        Cell cell = worksheet.Cells["A1"];
        cell.Value = "Hello Aspose";

        // Obtain the RichTextPortion (FontSetting) for the word "Aspose"
        // Start index 6 (zero‑based) and length 6 characters
        FontSetting richTextPortion = cell.Characters(6, 6);

        // Change the font size of this portion to 12 points
        richTextPortion.Font.Size = 12;

        // Save the workbook
        workbook.Save("RichTextPortionFontSize.xlsx");
    }
}