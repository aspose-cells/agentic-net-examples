// Title: C# – Insert Multi‑Line Text into an Aspose.Cells ListObject Table Cell with PutCellValue
// Description: Shows how to create a workbook, add a ListObject (Excel table), and use ListObject.PutCellValue(rowOffset, columnOffset, value) to write a string that contains newline characters (e.g., "Line 1\nLine 2\nLine 3") into a table cell so the text displays on separate lines in Excel.
// Keywords: Aspose.Cells | ListObject | PutCellValue | C# | .NET | multi line cell | newline characters | line break in Excel table | wrap text | Excel table cell value
// Common Searches: Aspose.Cells C# insert multi line text into table cell | ListObject.PutCellValue newline characters | How to add line breaks to an Excel table cell with Aspose.Cells | C# Aspose.Cells write multi‑line string to ListObject cell | Excel table cell wrap text using Aspose.Cells
// Developer Intent: Write a multi‑line string into a specific cell of an Aspose.Cells ListObject (Excel table) using PutCellValue.
// Use Cases: Display address or notes that require line breaks inside a single table cell. | Generate reports where a description column contains bullet points or paragraph text. | Preserve formatting when exporting database fields with embedded newline characters to Excel.
// AI Prompts: Provide a C# example that uses ListObject.PutCellValue to insert a string with \n line breaks into an Excel table cell and ensures the cell wraps text. | Show how to enable text wrapping for a ListObject cell after inserting multi‑line content with Aspose.Cells. | Explain the difference between PutCellValue and PutValue when adding newline characters to a table cell in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to create a workbook, add a ListObject (Excel table), and use ListObject.PutCellValue(rowOffset, columnOffset, value) to write a string that contains newline characters (e.g., "Line 1\nLine 2\nLine 3") into a table cell so the text displays on separate lines in Excel.
class InsertMultiLineString
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row and a sample data row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Description");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Initial");

        // Create a ListObject (table) that includes the header and data rows
        int tableIndex = sheet.ListObjects.Add("A1", "B2", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Define a multi‑line string using newline characters
        string multiLine = "Line 1\nLine 2\nLine 3";

        // Insert the multi‑line string into the cell at row offset 1, column offset 1 of the table
        // (this corresponds to cell B2 in the worksheet)
        table.PutCellValue(1, 1, multiLine);

        // Save the workbook to a file
        workbook.Save("MultiLineTableCell.xlsx", SaveFormat.Xlsx);
    }
}
