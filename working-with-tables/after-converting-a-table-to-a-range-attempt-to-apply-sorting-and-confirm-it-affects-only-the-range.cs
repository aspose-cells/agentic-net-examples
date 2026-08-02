// Title: Sort a converted table range while preserving other cells – Aspose.Cells for .NET
// Description: Demonstrates how to convert a ListObject to a normal range, apply DataSorter to the range, and verify that cells outside the sorted area remain unchanged. The example creates a workbook, fills data, adds a table on A1:B5, converts it, sorts by the ID column, and checks the value in D2 before saving.
// Keywords: Aspose.Cells sort range after ConvertToRange | C# DataSorter CellArea | ListObject to range Aspose.Cells | preserve external cells during sort | Aspose.Cells sorting with headers
// Common Searches: Aspose.Cells sort only a specific range after converting a table | DataSorter.Sort affect cells outside CellArea | Convert ListObject to range and sort C# | keep side columns unchanged when sorting Aspose.Cells
// Developer Intent: Apply sorting to the cells that were part of a table after it has been converted to a regular range, ensuring that any other worksheet data stays untouched.
// Use Cases: Reorder rows of a data block that originated from a ListObject without disturbing adjacent summary columns. | Generate reports where only a defined area (e.g., A1:B5) needs sorting while other sections remain static. | Validate that a cell outside the sorted area (such as D2) retains its original value after the operation.
// AI Prompts: Write C# code using Aspose.Cells to convert a ListObject to a range and sort it by the first column without changing other cells. | Explain how DataSorter.Sort works with a CellArea and how to confirm that external cells are unaffected. | Provide a unit‑test snippet that asserts cell D2 still contains 999 after sorting the converted range.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to convert a ListObject to a normal range, apply DataSorter to the range, and verify that cells outside the sorted area remain unchanged. The example creates a workbook, fills data, adds a table on A1:B5, converts it, sorts by the ID column, and checks the value in D2 before saving.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate data with a header row and several data rows
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue(3);
        cells["B2"].PutValue(30);
        cells["A3"].PutValue(1);
        cells["B3"].PutValue(10);
        cells["A4"].PutValue(4);
        cells["B4"].PutValue(40);
        cells["A5"].PutValue(2);
        cells["B5"].PutValue(20);

        // Add some data outside the intended range to prove sorting is limited
        cells["D1"].PutValue("Outside");
        cells["D2"].PutValue(999);

        // Create a table (ListObject) that covers the data range A1:B5
        int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "SampleTable";

        // Convert the table back to a normal range
        table.ConvertToRange();

        // Define the area that should be sorted (including the header row)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            StartColumn = 0,   // Column A
            EndRow = 4,        // Row 5
            EndColumn = 1      // Column B
        };

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;                     // First row contains headers
        sorter.AddKey(0, SortOrder.Ascending);        // Sort by the first column (ID)

        // Perform the sort on the defined area
        sorter.Sort(worksheet.Cells, sortArea);

        // Optional: output a cell outside the sorted area to confirm it stayed unchanged
        Console.WriteLine("Value in D2 (should be 999): " + cells["D2"].IntValue);

        // Save the workbook
        workbook.Save("SortedAfterConvertToRange.xlsx");
    }
}
