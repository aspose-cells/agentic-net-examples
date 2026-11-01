// Title: Insert a hyperlink into a specific ListObject cell using ListObject.PutCellValue in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add a hyperlink formula to a table cell using ListObject.PutCellValue | How to set an Excel HYPERLINK in a ListObject column with Aspose.Cells .NET | Aspose.Cells example for inserting a clickable link into a table row programmatically | Use PutCellValue to write a HYPERLINK formula into a ListObject cell in C# | Create an Excel table and add a hyperlink to a cell using the Aspose.Cells API
// Tags: ListObject.PutCellValue with HYPERLINK formula | Aspose.Cells add hyperlink to table cell | C# insert Excel hyperlink into ListObject | Aspose.Cells table cell value assignment | Excel table hyperlink using Aspose.Cells .NET | programmatic hyperlink insertion in Aspose.Cells ListObject

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a new workbook, defines the range A1:B3 as a ListObject named "MyTable", and uses ListObject.PutCellValue to place an Excel HYPERLINK formula (=HYPERLINK("https://www.example.com","Example Site")) into the second row, second column of the table. The workbook is then saved as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header and sample data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define the range that will become the ListObject (table)
            CellArea tableArea = new CellArea
            {
                StartRow = 0,    // Row 1 (zero‑based)
                StartColumn = 0, // Column A
                EndRow = 2,      // Row 3
                EndColumn = 1    // Column B
            };

            // Convert CellArea to an address string (e.g., "A1:B3")
            string startCell = CellsHelper.CellIndexToName(tableArea.StartRow, tableArea.StartColumn);
            string endCell = CellsHelper.CellIndexToName(tableArea.EndRow, tableArea.EndColumn);
            string source = $"{startCell}:{endCell}";

            // Add the ListObject to the worksheet (provide a name for the table)
            int listObjectIndex = sheet.ListObjects.Add("MyTable", source, true);
            ListObject table = sheet.ListObjects[listObjectIndex];

            // Insert a hyperlink using the Excel HYPERLINK formula
            string hyperlinkFormula = "=HYPERLINK(\"https://www.example.com\",\"Example Site\")";
            table.PutCellValue(1, 1, hyperlinkFormula); // Row index 1 = second row, Column index 1 = second column

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
