// Title: Insert a multi‑line string into an Aspose.Cells ListObject table cell with text wrapping in C#
// Common Searches: C# Aspose.Cells put multiline text into a ListObject table cell | Insert line breaks into Excel table cell using ListObject.PutCellValue | Enable text wrap for multi‑line cell values in Aspose.Cells tables | Aspose.Cells ListObject.PutCellValue with \r\n characters in .NET | Add multi‑line string to Excel table data range using Aspose.Cells
// Tags: ListObject.PutCellValue multiline | Aspose.Cells table cell newline | text wrapping style Aspose.Cells | C# Excel table cell line breaks | Aspose.Cells data range styling

using Aspose.Cells;

// The example creates a workbook, defines a ListObject covering A1:A2, builds a string containing \r\n line breaks, inserts the multi‑line string into the first data cell of the table with ListObject.PutCellValue, applies a wrapped‑text style to the cell, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Populate header and a single data row to define the table range
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["A2"].PutValue("Data");

        // Add a ListObject (table) that covers the range A1:A2
        int listObjectIndex = worksheet.ListObjects.Add(0, 0, 1, 0, true);
        var listObject = worksheet.ListObjects[listObjectIndex];
        listObject.DisplayName = "MyTable";

        // Multi‑line string with newline characters
        string multiLine = "First line\r\nSecond line\r\nThird line";

        // Insert the multi‑line string into the first data row, first column of the table
        // Row index 0 refers to the first data row (excluding the header)
        listObject.PutCellValue(0, 0, multiLine);

        // Enable text wrapping so the newlines are visible in Excel
        var cell = listObject.DataRange[0, 0];
        var style = cell.GetStyle();
        style.IsTextWrapped = true;
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
