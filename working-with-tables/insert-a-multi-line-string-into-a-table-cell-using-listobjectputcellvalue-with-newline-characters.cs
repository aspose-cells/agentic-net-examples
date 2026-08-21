// Title: Insert Multi‑Line Text into an Aspose.Cells ListObject Table Cell with PutCellValue (C#)
// Description: Creates a workbook, defines a table spanning A1:B3, builds a string that includes "\r\n" line breaks, and calls ListObject.PutCellValue(rowOffset, columnOffset, text) to write the wrapped content into the second data row, second column before saving as MultiLineTableCell.xlsx.
// Keywords: Aspose.Cells ListObject PutCellValue | C# multiline cell | Excel table line break | programmatic text wrap | insert newline in table cell
// Common Searches: Aspose.Cells how to add line break in table cell | PutCellValue multiline C# example | Set wrapped text in ListObject cell | Insert address with newlines using Aspose.Cells | Create Excel table with multi‑line description column
// Developer Intent: Place a string that contains line‑break characters into a targeted cell of a ListObject table.
// Use Cases: Populate a description column with bullet‑point style entries that span several lines. | Store address or notes fields as wrapped text inside table cells for clearer reports. | Generate an Excel worksheet where multi‑line strings improve readability of tabular data.
// AI Prompts: Show how to enable text wrapping for a ListObject cell after inserting a multi‑line string with PutCellValue. | Provide C# code that writes a multi‑line value into a table cell and automatically adjusts the row height for proper display. | Explain the effect of using "\n" versus "\r\n" when passing newline characters to ListObject.PutCellValue.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, defines a table spanning A1:B3, builds a string that includes "\r\n" line breaks, and calls ListObject.PutCellValue(rowOffset, columnOffset, text) to write the wrapped content into the second data row, second column before saving as MultiLineTableCell.xlsx.
class InsertMultiLineStringIntoTableCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row and some sample data for the table
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Description");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Item 1");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Item 2");

        // Create a ListObject (table) that spans the range A1:B3
        int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Define a multi‑line string using newline characters
        string multiLineText = "First line\r\nSecond line\r\nThird line";

        // Insert the multi‑line string into the second data row, second column of the table
        // Row offset = 1 (zero‑based, so row 2 of the table), column offset = 1 (second column)
        table.PutCellValue(1, 1, multiLineText);

        // Save the workbook to a file
        workbook.Save("MultiLineTableCell.xlsx", SaveFormat.Xlsx);
    }
}
