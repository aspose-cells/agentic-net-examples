using System;
using Aspose.Cells;

class SortSpreadsheet
{
    static void Main()
    {
        // Load the existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;          // First row contains column headers
        sorter.Key1 = 1;                    // Sort by the second column (index 1)
        sorter.Order1 = SortOrder.Ascending; // Ascending order

        // Define the range to be sorted (including headers)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = worksheet.Cells.MaxDataRow,       // Last row with data
            EndColumn = worksheet.Cells.MaxDataColumn   // Last column with data
        };

        // Perform the sort; formatting and data types are preserved automatically
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook to a new file
        workbook.Save("sorted_output.xlsx");
    }
}