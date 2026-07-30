// Title: C# – Add a Hyperlink to a ListObject Cell with ListObject.PutCellValue in Aspose.Cells
// Description: Shows how to create a workbook, define a ListObject (Excel table), set cell text via ListObject.PutCellValue, obtain the absolute cell address, and attach a hyperlink using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | ListObject | PutCellValue | hyperlink | Excel table | add hyperlink to table cell | Aspose.Cells .NET example | programmatic Excel hyperlink | GitHub Aspose.Cells sample
// Common Searches: how to add hyperlink to a cell in Aspose.Cells ListObject | ListObject.PutCellValue hyperlink C# | Aspose.Cells add hyperlink to table column | C# Aspose.Cells create hyperlink in Excel table | Aspose.Cells hyperlink example
// Developer Intent: Insert display text and attach a URL to a specific cell inside an Aspose.Cells ListObject table.
// Use Cases: Generate a report where each row contains an ID and a clickable link with custom text. | Automate a documentation index workbook that stores external URLs in a table column. | Populate a data‑driven table where the hyperlink text is set programmatically and the URL is added afterward.
// AI Prompts: Write C# code using Aspose.Cells to set text in a ListObject cell with PutCellValue and then add a hyperlink to the same cell. | Show how to retrieve the absolute address of a ListObject cell after PutCellValue and use it to create a hyperlink. | Explain how to add different URLs with unique display texts to multiple rows of an Aspose.Cells table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsHyperlinkInTable
{
    // Shows how to create a workbook, define a ListObject (Excel table), set cell text via ListObject.PutCellValue, obtain the absolute cell address, and attach a hyperlink using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Link");

            // Add some sample data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(3);

            // Define the range that will become a ListObject (table)
            // Table covers A1:B4 (including header and three data rows)
            int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Insert a hyperlink display text into the second column of the second data row
            // Row offset is 1 (first data row after header), column offset is 1 (second column)
            table.PutCellValue(1, 1, "Visit Aspose");

            // Determine the absolute cell address where the value was placed
            int targetRow = table.StartRow + 1;      // StartRow is zero‑based index of the first row of the table
            int targetColumn = table.StartColumn + 1; // StartColumn is zero‑based index of the first column of the table
            string cellName = sheet.Cells[targetRow, targetColumn].Name; // e.g., "B2"

            // Add a hyperlink to the same cell with the desired URL
            sheet.Hyperlinks.Add(cellName, 1, 1, "https://www.aspose.com");

            // Optionally set the display text of the hyperlink (overwrites the cell value if needed)
            // Here we keep the previously set text, so no extra action is required.

            // Save the workbook
            workbook.Save("TableWithHyperlink.xlsx", SaveFormat.Xlsx);
        }
    }
}
