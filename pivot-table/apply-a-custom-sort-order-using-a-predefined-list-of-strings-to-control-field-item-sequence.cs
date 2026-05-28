using System;
using Aspose.Cells;

class CustomSortDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in column A
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Medium");
        worksheet.Cells["A3"].PutValue("Low");
        worksheet.Cells["A4"].PutValue("High");
        worksheet.Cells["A5"].PutValue("Critical");

        // Define the custom sort order (desired sequence)
        string[] customList = new string[] { "Critical", "High", "Medium", "Low" };

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // First row contains headers
        // Add a sort key for column A (index 0) using the custom list
        sorter.AddKey(0, SortOrder.Ascending, customList);

        // Perform the sort on the defined range (including header)
        sorter.Sort(worksheet.Cells, CellArea.CreateCellArea("A1", "A5"));

        // Save the workbook with the custom-sorted data
        workbook.Save("CustomSorted.xlsx");
    }
}