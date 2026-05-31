using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add header and unsorted month data in column C
        cells["C1"].PutValue("Month");
        cells["C2"].PutValue("March");
        cells["C3"].PutValue("January");
        cells["C4"].PutValue("December");
        cells["C5"].PutValue("July");
        cells["C6"].PutValue("May");

        // Add some additional columns to demonstrate full‑row sorting
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue(1); cells["B2"].PutValue(10);
        cells["A3"].PutValue(2); cells["B3"].PutValue(20);
        cells["A4"].PutValue(3); cells["B4"].PutValue(30);
        cells["A5"].PutValue(4); cells["B5"].PutValue(40);
        cells["A6"].PutValue(5); cells["B6"].PutValue(50);

        // Configure the DataSorter with a custom month list
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // First row contains headers
        string monthList = "January,February,March,April,May,June,July,August,September,October,November,December";
        sorter.AddKey(2, SortOrder.Ascending, monthList); // Column C index = 2

        // Define the range to sort (columns A‑C, rows 1‑6)
        CellArea area = CellArea.CreateCellArea("A1", "C6");

        // Perform the sort
        sorter.Sort(sheet.Cells, area);

        // Save the sorted workbook
        workbook.Save("SortedMonths.xlsx");
    }
}