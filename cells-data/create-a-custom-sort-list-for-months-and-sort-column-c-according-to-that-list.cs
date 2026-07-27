using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header row
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["C1"].PutValue("Month");

        // Add sample data (months are in column C)
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["C2"].PutValue("March");

        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["C3"].PutValue("January");

        worksheet.Cells["A4"].PutValue(3);
        worksheet.Cells["B4"].PutValue(150);
        worksheet.Cells["C4"].PutValue("December");

        worksheet.Cells["A5"].PutValue(4);
        worksheet.Cells["B5"].PutValue(120);
        worksheet.Cells["C5"].PutValue("July");

        // Define a custom sort list for months
        string monthCustomList = "January,February,March,April,May,June,July,August,September,October,November,December";

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // First row contains headers
        // Add sorting key for column C (index 2) using the custom month list
        sorter.AddKey(2, SortOrder.Ascending, monthCustomList);

        // Define the range to sort (including the header row)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 4,
            EndColumn = 2
        };

        // Perform the sort
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the workbook
        workbook.Save("SortedMonths.xlsx");
    }
}