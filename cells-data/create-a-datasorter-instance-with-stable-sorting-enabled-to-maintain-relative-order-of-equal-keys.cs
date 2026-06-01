using Aspose.Cells;
using System;

class StableSortingExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data with duplicate keys to demonstrate stable sorting
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("A");
        cells["B4"].PutValue(15);
        cells["A5"].PutValue("B");
        cells["B5"].PutValue(5);
        cells["A6"].PutValue("A");
        cells["B6"].PutValue(12);

        // Obtain the DataSorter instance (Aspose.Cells uses a stable sort algorithm by default)
        DataSorter sorter = workbook.DataSorter;

        // Define sorting keys
        sorter.Key1 = 0;                     // First key: Column A (Category)
        sorter.Order1 = SortOrder.Ascending; // Ascending order for the first key

        sorter.Key2 = 1;                     // Second key: Column B (Value)
        sorter.Order2 = SortOrder.Descending; // Descending order for the second key

        // Define the range to sort (including headers)
        CellArea range = CellArea.CreateCellArea("A1", "B6");

        // Perform the sort; equal keys will retain their original relative order
        sorter.Sort(cells, range);

        // Save the sorted workbook
        workbook.Save("StableSortedOutput.xlsx");
    }
}