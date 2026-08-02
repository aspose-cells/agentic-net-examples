// Title: C# – Add a Boolean to a Specific Row in an Aspose.Cells ListObject Table with PutCellValue
// Description: Creates a workbook, defines a ListObject over A1:B3, and uses ListObject.PutCellValue with zero‑based row and column offsets to insert a new data row and set the Boolean Flag column to true, then saves the file as an XLSX document.
// Keywords: Aspose.Cells ListObject PutCellValue | C# add boolean to table row | Aspose.Cells table row offset | set cell value in ListObject | Aspose.Cells boolean column example
// Common Searches: Aspose.Cells PutCellValue boolean C# | how to add a new row to ListObject and set flag | zero based row offset ListObject PutCellValue | C# Aspose.Cells set true/false in table column | insert boolean value into Aspose.Cells table
// Developer Intent: Insert a Boolean value into a designated row and column of a ListObject table using PutCellValue.
// Use Cases: Automatically assign a true/false status flag when appending new records to a spreadsheet table. | Update a Boolean field in an existing table row based on calculation results or user input. | Generate dynamic reports where each row includes a Boolean indicator for filtering or conditional formatting.
// AI Prompts: Write C# code that adds a new row to an Aspose.Cells ListObject and sets a Boolean value in the second column using PutCellValue. | Explain the concept of zero‑based row and column offsets in ListObject.PutCellValue with a code example. | Show how to modify an existing Boolean cell in a ListObject without adding a new row, using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, defines a ListObject over A1:B3, and uses ListObject.PutCellValue with zero‑based row and column offsets to insert a new data row and set the Boolean Flag column to true, then saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row and some initial data to form a table
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Flag");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue(false);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue(false);

        // Create a ListObject (table) that includes the range A1:B3, with a header row
        int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Add a new row to the table (row offset 2) and set a boolean value in column 1
        // Row offsets are zero‑based relative to the first data row (excluding the header)
        table.PutCellValue(2, 0, 3);      // Set ID = 3 in the new row
        table.PutCellValue(2, 1, true);  // Set boolean Flag = true in the new row

        // Save the workbook to a file
        workbook.Save("ListObjectBooleanDemo.xlsx", SaveFormat.Xlsx);
    }
}
