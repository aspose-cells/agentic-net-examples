using System;
using Aspose.Cells;

namespace AsposeCellsCustomTextSort
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a header and textual values
            cells["A1"].PutValue("Category");   // Header
            cells["A2"].PutValue("Banana");
            cells["A3"].PutValue("Apple");
            cells["A4"].PutValue("Cherry");
            cells["A5"].PutValue("Date");

            // Configure the DataSorter to sort column A in descending order
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                     // First row is a header
            sorter.AddKey(0, SortOrder.Descending);       // Column index 0 (A), descending

            // Define the range that includes the header and data rows
            CellArea sortArea = CellArea.CreateCellArea("A1", "A5");

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // Save the workbook (lifecycle: save)
            workbook.Save("CustomTextDescendingSort.xlsx");
        }
    }
}