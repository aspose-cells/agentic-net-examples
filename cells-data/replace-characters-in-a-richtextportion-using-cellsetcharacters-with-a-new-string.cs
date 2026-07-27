using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates replacing a portion of rich‑text in a cell using Cell.SetCharacters
    class ReplaceRichTextPortionDemo
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Access a cell and set an initial rich‑text value
            Cell cell = sheet.Cells["A1"];
            cell.Value = "Hello World!";

            // 3. Define the new text that will replace the original portion ("World")
            string newPortion = "Aspose";

            // 4. Build the final string for the cell (replace "World" with "Aspose")
            //    Original: "Hello World!"
            //    Result:   "Hello Aspose!"
            string finalText = cell.StringValue.Replace("World", newPortion);
            cell.Value = finalText; // update cell content

            // 5. Create a FontSetting for the replaced portion to preserve formatting
            //    The replaced text starts at index 6 (after "Hello ") and has length equal to newPortion.Length
            FontSetting replacedSetting = new FontSetting(6, newPortion.Length, workbook.Worksheets);
            // Example formatting: make the new portion bold and blue
            replacedSetting.Font.IsBold = true;
            replacedSetting.Font.Color = Color.Blue;

            // 6. Apply the FontSetting array to the cell using SetCharacters
            FontSetting[] settings = new FontSetting[] { replacedSetting };
            cell.SetCharacters(settings);

            // 7. Save the workbook (output file)
            workbook.Save("ReplaceRichTextPortionDemo.xlsx");
        }
    }
}