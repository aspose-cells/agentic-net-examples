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

        // Add sample data (first row is a header)
        cells["A1"].PutValue("Name");
        cells["A2"].PutValue("John");
        cells["A3"].PutValue("Alice");
        cells["A4"].PutValue("Bob");

        // Configure the DataSorter to sort by column A (index 0) in ascending order
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;   // first row contains headers
        sorter.Key1 = 0;            // column A
        sorter.Order1 = SortOrder.Ascending;

        // Define the range to sort (including the header row)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 3,               // last data row index
            EndColumn = cells.MaxDataColumn
        };

        // Perform the sort
        sorter.Sort(cells, area);

        // Save the sorted workbook
        workbook.Save("SortedByColumnA.xlsx");
    }
}