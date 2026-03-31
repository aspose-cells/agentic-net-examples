using System;
using Aspose.Cells;

namespace AsposeCellsSortingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including a header row)
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Laptop");
            cells["B2"].PutValue("Electronics");
            cells["C2"].PutValue(1200);

            cells["A3"].PutValue("Desk");
            cells["B3"].PutValue("Furniture");
            cells["C3"].PutValue(300);

            cells["A4"].PutValue("Smartphone");
            cells["B4"].PutValue("Electronics");
            cells["C4"].PutValue(800);

            cells["A5"].PutValue("Chair");
            cells["B5"].PutValue("Furniture");
            cells["C5"].PutValue(150);

            // Configure the DataSorter to sort by Price (column C) in ascending order
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;               // First row contains headers
            sorter.Key1 = 2;                         // Column index for "Price" (C = 2)
            sorter.Order1 = SortOrder.Ascending;    // Ascending order

            // Define the range to be sorted (including headers)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,      // Header row
                StartColumn = 0,   // Column A
                EndRow = 4,        // Last data row (zero‑based index)
                EndColumn = 2      // Column C
            };

            // Perform the sort using Excel's native sorting engine
            sorter.Sort(cells, sortArea);

            // Save the sorted workbook to an XLSX file
            workbook.Save("SortedProducts.xlsx");
        }
    }
}