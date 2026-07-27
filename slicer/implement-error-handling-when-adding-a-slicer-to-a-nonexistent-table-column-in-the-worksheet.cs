using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using System;

class SlicerErrorHandlingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the table
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");

        // Add a table that covers the data range A1:B3
        int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Index of the column we want to use for the slicer (zero‑based)
        // Valid indices for this table are 0 (ID) and 1 (Name)
        int targetColumnIndex = 2; // Intentionally invalid to demonstrate error handling

        try
        {
            // Verify that the requested column exists in the table
            if (targetColumnIndex < 0 || targetColumnIndex >= table.ListColumns.Count)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(targetColumnIndex),
                    $"Column index {targetColumnIndex} is out of range. The table contains {table.ListColumns.Count} columns.");
            }

            // Column exists – add the slicer at row 5, column 0 (cell A5)
            SlicerCollection slicers = sheet.Slicers;
            slicers.Add(table, table.ListColumns[targetColumnIndex], 5, 0);
            Console.WriteLine("Slicer added successfully.");
        }
        catch (Exception ex)
        {
            // Handle the error (log, display, etc.)
            Console.WriteLine($"Error adding slicer: {ex.Message}");
        }

        // Save the workbook
        workbook.Save("SlicerErrorHandlingDemo.xlsx");
    }
}