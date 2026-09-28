// Title: How to convert an Aspose.Cells ListObject to a Range and copy it to another worksheet at a specific cell using C#
// AI Prompts: Write C# code that creates a ListObject, extracts its cells as a Range, and copies that range to a second worksheet beginning at cell C5 with Aspose.Cells. | Demonstrate how to use the Aspose.Cells Range.Copy method to move a table from one sheet to another, including the creation of source and target ranges. | Provide a full example that builds a workbook, adds a table, defines a destination range on another sheet, and copies the table using Aspose.Cells for .NET.
// Common Searches: asp.net how to copy a table from one worksheet to another using Aspose.Cells | example of Range.Copy to paste a ListObject at cell C5 in a different sheet | c# aspose.cells duplicate an Excel table on another worksheet | using Aspose.Cells to move a table range to a specific cell location | copying a ListObject to a destination range with Aspose.Cells C#
// Tags: convert ListObject to Range Aspose.Cells | copy range between worksheets C# | Aspose.Cells Range.Copy usage | duplicate Excel table programmatically | copy table to specific cell Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds a ListObject (table) on the first sheet, converts the table to a Range, defines a destination range starting at cell C5 on a second sheet, copies the source range to the destination using Range.Copy, and saves the file as TableCopy.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it
            Worksheet wsSource = workbook.Worksheets[0];
            wsSource.Name = "Source";

            // Populate sample data for the table
            wsSource.Cells["A1"].PutValue("ID");
            wsSource.Cells["B1"].PutValue("Name");
            wsSource.Cells["A2"].PutValue(1);
            wsSource.Cells["B2"].PutValue("Alice");
            wsSource.Cells["A3"].PutValue(2);
            wsSource.Cells["B3"].PutValue("Bob");

            // Define the area of the table (A1:B3)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // rows including header
            int totalColumns = 2;      // columns

            // Add a ListObject (table) over the defined range; hasHeaders = true
            int tableIndex = wsSource.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = wsSource.ListObjects[tableIndex];
            // Set display name for the table (Name property is not available in this API version)
            table.DisplayName = "MyTable";

            // Convert the table to a Range (including header)
            Aspose.Cells.Range tableRange = wsSource.Cells.CreateRange(
                firstRow,
                firstColumn,
                totalRows,
                totalColumns);

            // Add a destination worksheet
            Worksheet wsDest = workbook.Worksheets.Add("Destination");

            // Define the destination range starting at cell C5 (row 5, column C)
            int destStartRow = 4;      // zero‑based index for row 5
            int destStartColumn = 2;   // zero‑based index for column C
            Aspose.Cells.Range destRange = wsDest.Cells.CreateRange(
                destStartRow,
                destStartColumn,
                totalRows,
                totalColumns);

            // Copy the source range to the destination range
            tableRange.Copy(destRange);

            // Save the workbook to a file
            workbook.Save("TableCopy.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
