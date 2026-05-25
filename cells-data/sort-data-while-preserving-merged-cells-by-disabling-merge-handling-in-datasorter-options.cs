using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class SortPreserveMergedCells
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // Set up sample data with a merged header row
            // -------------------------------------------------
            // Merge cells A1:C1 to create a header that spans three columns
            cells.Merge(0, 0, 1, 3);
            cells[0, 0].PutValue("Sales Report"); // Header text

            // Add column titles (these will be sorted)
            cells["A2"].PutValue("Region");
            cells["B2"].PutValue("Quarter");
            cells["C2"].PutValue("Amount");

            // Add some data rows
            cells["A3"].PutValue("East");
            cells["B3"].PutValue(1);
            cells["C3"].PutValue(5000);

            cells["A4"].PutValue("West");
            cells["B4"].PutValue(2);
            cells["C4"].PutValue(3000);

            cells["A5"].PutValue("North");
            cells["B5"].PutValue(1);
            cells["C5"].PutValue(7000);

            cells["A6"].PutValue("South");
            cells["B6"].PutValue(2);
            cells["C6"].PutValue(2000);

            // -------------------------------------------------
            // Configure the DataSorter
            // -------------------------------------------------
            DataSorter sorter = workbook.DataSorter;

            // The data has a header row (the merged header plus column titles)
            sorter.HasHeaders = true;

            // Sort by the "Amount" column (index 2) in descending order
            sorter.AddKey(2, SortOrder.Descending);

            // Define the area to sort:
            // StartRow = 1 (row 2 in Excel, includes column titles)
            // EndRow   = 5 (row 6 in Excel, last data row)
            // StartColumn = 0 (A), EndColumn = 2 (C)
            CellArea sortArea = new CellArea
            {
                StartRow = 1,
                EndRow = 5,
                StartColumn = 0,
                EndColumn = 2
            };

            // -------------------------------------------------
            // Perform the sort
            // -------------------------------------------------
            // By default DataSorter does not alter merged cells.
            // No special option is required to "disable merge handling".
            sorter.Sort(cells, sortArea);

            // -------------------------------------------------
            // Verify that the merged header is still intact
            // -------------------------------------------------
            Console.WriteLine("Merged header after sort: " + cells[0, 0].StringValue);
            Console.WriteLine("Is the header still merged? " + (cells.MergedCells.Count > 0));

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("SortedPreserveMerged.xlsx");
        }
    }
}