using System;
using Aspose.Cells;

class AddBoldRichTextPortion
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the target cell and set initial plain text
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Hello World");

        // Text to be added as a new rich text portion
        string newText = " BoldPart";

        // Insert the new text at the end of the existing cell value
        int insertIndex = cell.StringValue.Length; // position after current text
        cell.InsertText(insertIndex, newText);

        // Apply bold formatting to the inserted portion
        cell.Characters(insertIndex, newText.Length).Font.IsBold = true;

        // Save the workbook
        workbook.Save("RichTextBold.xlsx");
    }
}