// Title: Convert Aspose.Cells ListObject to a Range and Sort Columns A‑B While Preserving Other Data (C#)
// Description: Creates a workbook, adds a ListObject (A1:C6), converts it to a normal range, defines a CellArea covering columns A and B, configures DataSorter with headers, sorts by the Category column, and saves the file. Column C (Info) and column D (outside data) remain unchanged, proving the sort affects only the selected range.
// Keywords: Aspose.Cells C# convert ListObject to range | Aspose.Cells DataSorter sort specific columns | preserve adjacent columns when sorting Aspose.Cells | ConvertToRange example Aspose.Cells | sort range with headers Aspose.Cells
// Common Searches: How to convert an Aspose.Cells table to a range and sort part of it | Sorting a range after removing a ListObject in Aspose.Cells | Aspose.Cells keep extra columns unchanged during sort | DataSorter sort only selected columns after ConvertToRange
// Developer Intent: Remove a ListObject while keeping the cell data, then sort only chosen columns without affecting other columns.
// Use Cases: Drop table formatting to apply custom sorting on specific columns only. | Use DataSorter.HasHeaders with a defined CellArea to sort a subset of data after a table is converted to a range. | Validate that columns outside the sort area, including extra columns originally inside the table, stay unchanged after sorting.
// AI Prompts: Show C# code that converts an Aspose.Cells ListObject to a normal range and sorts only columns A and B, leaving column C and any outside columns untouched. | Generate an Aspose.Cells example that defines a CellArea for sorting, sets HasHeaders, adds a sort key for the Category column, and saves the workbook. | Explain how Aspose.Cells DataSorter works on a range after ConvertToRange and how to verify that non‑sorted columns are not modified.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds a ListObject (A1:C6), converts it to a normal range, defines a CellArea covering columns A and B, configures DataSorter with headers, sorts by the Category column, and saves the file. Column C (Info) and column D (outside data) remain unchanged, proving the sort affects only the selected range.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate data with a header row (A1:C1) and some extra data outside the table (column D)
            cells["A1"].PutValue("Category");   // Header
            cells["B1"].PutValue("Value");      // Header
            cells["C1"].PutValue("Info");       // Extra column inside the future table range (will stay untouched by sort)

            string[] categories = { "A", "B", "A", "B", "A" };
            int[] values = { 10, 20, 30, 40, 50 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]);               // Column A (Category)
                cells[i + 1, 1].PutValue(values[i]);                  // Column B (Value)
                cells[i + 1, 2].PutValue($"Row{i + 1}");              // Column C (Info) – not part of the sort area
            }

            // Data outside the table to verify it remains unchanged after sorting
            cells["D1"].PutValue("OutsideHeader");
            cells["D2"].PutValue("OutsideData");

            // Create a ListObject (table) that covers the header and data rows (A1:C6)
            int tableIndex = worksheet.ListObjects.Add("A1", $"C{categories.Length + 1}", true);
            ListObject table = worksheet.ListObjects[tableIndex];
            // The ShowHeaders property is not required; headers are already recognized via the third argument above.

            // Convert the table back to a normal range; the table object is removed but the cells stay intact
            table.ConvertToRange();

            // Define the area to sort: only the data columns A and B (including the header row)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,                     // Row 1 (zero‑based)
                StartColumn = 0,                  // Column A
                EndRow = categories.Length,       // Last data row (row index 5)
                EndColumn = 1                     // Column B
            };

            // Configure the DataSorter to treat the first row as headers and sort by the "Category" column (index 0)
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;
            sorter.AddKey(0, SortOrder.Ascending); // Sort ascending by Category

            // Perform the sort on the defined range
            sorter.Sort(worksheet.Cells, sortArea);

            // Save the workbook – the sorted range should reflect the new order,
            // while column C (Info) and column D (outside data) remain unchanged.
            workbook.Save("SortedAfterConvertToRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
