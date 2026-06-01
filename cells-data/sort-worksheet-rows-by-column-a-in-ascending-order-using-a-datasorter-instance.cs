using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // -------------------------------------------------
        // Sample data (you can replace this with your own)
        // -------------------------------------------------
        cells["A1"].PutValue("Charlie");
        cells["B1"].PutValue(30);
        cells["A2"].PutValue("Alice");
        cells["B2"].PutValue(25);
        cells["A3"].PutValue("Bob");
        cells["B3"].PutValue(28);
        // -------------------------------------------------

        // Get the DataSorter object from the workbook
        DataSorter sorter = workbook.DataSorter;

        // Set the sort key to column A (index 0) and order to ascending
        sorter.Key1 = 0;                 // Column A
        sorter.Order1 = SortOrder.Ascending;

        // Define the range to sort (all used rows and columns)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = cells.MaxDataRow,
            EndColumn = cells.MaxDataColumn
        };

        // Perform the sort
        sorter.Sort(cells, sortArea);

        // Save the sorted workbook
        workbook.Save("SortedByColumnA.xlsx");
    }
}