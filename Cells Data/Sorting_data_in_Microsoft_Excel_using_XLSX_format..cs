using System;
using Aspose.Cells;

namespace AsposeCellsSortingDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (including a header row)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("Score");

            cells["A2"].PutValue("John");
            cells["B2"].PutValue(25);
            cells["C2"].PutValue(85);

            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(30);
            cells["C3"].PutValue(92);

            cells["A4"].PutValue("Bob");
            cells["B4"].PutValue(22);
            cells["C4"].PutValue(78);

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;               // First row contains headers
            sorter.Key1 = 2;                         // Sort by the third column (Score)
            sorter.Order1 = SortOrder.Descending;    // Highest scores first

            // Define the range to be sorted (including headers)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,      // Row 0 (A1) – header row
                StartColumn = 0,   // Column A
                EndRow = 4,        // Row 4 (last data row)
                EndColumn = 2      // Column C
            };

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // Save the sorted workbook
            workbook.Save("SortedData.xlsx", SaveFormat.Xlsx);
        }
    }
}