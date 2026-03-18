using System;
using Aspose.Cells;

namespace AsposeCellsSortingDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputData.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Configure the DataSorter
            // ------------------------------------------------------------
            DataSorter sorter = workbook.DataSorter;

            // Assume the first row contains headers
            sorter.HasHeaders = true;

            // Primary sort key: column B (index 1) – e.g., "Age"
            sorter.Key1 = 1;
            sorter.Order1 = SortOrder.Ascending;

            // Secondary sort key: column A (index 0) – e.g., "Name"
            sorter.Key2 = 0;
            sorter.Order2 = SortOrder.Descending;

            // ------------------------------------------------------------
            // Define the range to be sorted (including headers)
            // ------------------------------------------------------------
            CellArea sortArea = new CellArea
            {
                StartRow = 0,                         // first row (header)
                StartColumn = 0,                      // first column
                EndRow = cells.MaxDataRow,            // last row with data
                EndColumn = cells.MaxDataColumn       // last column with data
            };

            // Perform the sort – cell formatting moves together with the data
            sorter.Sort(cells, sortArea);

            // Save the sorted workbook (replace with your desired output path)
            workbook.Save("SortedData.xlsx");
        }
    }
}