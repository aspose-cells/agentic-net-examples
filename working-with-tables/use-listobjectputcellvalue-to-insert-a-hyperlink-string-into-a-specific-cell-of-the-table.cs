// Title: Add a Hyperlink to a ListObject Cell Using ListObject.PutCellValue in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a ListObject (table), set the display text of a specific table cell with ListObject.PutCellValue, compute its absolute address, and attach a hyperlink via the Hyperlinks collection. The file is saved as an XLSX document.
// Keywords: Aspose.Cells | ListObject.PutCellValue | C# hyperlink table cell | Excel table hyperlink Aspose | Add hyperlink to ListObject | Aspose.Cells .NET example | hyperlink cell address
// Common Searches: Aspose.Cells add hyperlink to ListObject cell | ListObject.PutCellValue hyperlink example | how to set hyperlink text in Aspose.Cells table | retrieve absolute address of ListObject cell | C# Aspose.Cells table hyperlink code
// Developer Intent: Insert a clickable link into a specific cell of an Aspose.Cells ListObject while defining the visible text.
// Use Cases: Export a report where each row includes a link to a related document. | Generate product catalogs with URLs embedded in table cells. | Create data extracts that provide quick navigation to external resources from within an Excel table.
// AI Prompts: Show C# code that adds a hyperlink to a ListObject cell using ListObject.PutCellValue and the Hyperlinks collection. | Explain how to calculate the absolute cell reference of a ListObject element for hyperlink insertion. | Provide a step‑by‑step guide to set display text and URL for a table cell in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a workbook, define a ListObject (table), set the display text of a specific table cell with ListObject.PutCellValue, compute its absolute address, and attach a hyperlink via the Hyperlinks collection. The file is saved as an XLSX document.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Prepare header row for the table
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Link");

        // Add sample data rows
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["A3"].PutValue(2);

        // Create a ListObject (table) that spans A1:B3
        int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Insert the display text for the hyperlink into the table cell
        // Row offset 1 (second row of the table), column offset 1 (second column)
        table.PutCellValue(1, 1, "Visit Aspose");

        // Determine the absolute cell address of the cell we just updated
        int targetRow = table.StartRow + 1;      // absolute row index
        int targetColumn = table.StartColumn + 1; // absolute column index
        string cellName = CellsHelper.CellIndexToName(targetRow, targetColumn);

        // Add a hyperlink to that cell
        sheet.Hyperlinks.Add(cellName, 1, 1, "https://www.aspose.com");

        // Optionally set the text that will be displayed for the hyperlink
        int hyperlinkIdx = sheet.Hyperlinks.Count - 1;
        sheet.Hyperlinks[hyperlinkIdx].TextToDisplay = "Visit Aspose";

        // Save the workbook
        workbook.Save("ListObjectHyperlinkDemo.xlsx", SaveFormat.Xlsx);
    }
}
