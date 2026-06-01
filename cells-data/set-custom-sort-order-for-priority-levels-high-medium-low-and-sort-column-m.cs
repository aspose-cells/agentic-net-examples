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

        // Sample data with a header in column M (index 12)
        cells["A1"].PutValue("ID");
        cells["M1"].PutValue("Priority");

        cells["A2"].PutValue(1);
        cells["M2"].PutValue("Low");

        cells["A3"].PutValue(2);
        cells["M3"].PutValue("High");

        cells["A4"].PutValue(3);
        cells["M4"].PutValue("Medium");

        cells["A5"].PutValue(4);
        cells["M5"].PutValue("Low");

        cells["A6"].PutValue(5);
        cells["M6"].PutValue("High");

        // Get the DataSorter object
        DataSorter sorter = workbook.DataSorter;

        // Indicate that the first row contains headers
        sorter.HasHeaders = true;

        // Define the custom sort order for the priority column
        string customList = "High,Medium,Low";

        // Add a sort key for column M (index 12) using the custom list
        sorter.AddKey(12, SortOrder.Ascending, customList);

        // Determine the used range of the worksheet
        int lastRow = cells.MaxDataRow;
        int lastCol = cells.MaxDataColumn;

        // Create a CellArea covering the entire used range
        CellArea range = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = lastRow,
            EndColumn = lastCol
        };

        // Perform the sort operation
        sorter.Sort(cells, range);

        // Save the sorted workbook
        workbook.Save("SortedPriority.xlsx");
    }
}