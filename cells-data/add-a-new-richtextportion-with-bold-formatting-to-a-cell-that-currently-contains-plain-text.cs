using System;
using Aspose.Cells;

class AddRichTextPortion
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];

        // Set initial plain text in the cell
        cell.PutValue("Hello World");

        // Text to add as a new rich‑text portion
        string newText = " BoldText";

        // Insert the new text at the end of the existing content
        int insertIndex = cell.StringValue.Length;
        cell.InsertText(insertIndex, newText);

        // Apply bold formatting to the inserted portion
        cell.Characters(insertIndex, newText.Length).Font.IsBold = true;

        // Save the workbook
        workbook.Save("RichTextPortion.xlsx");
    }
}