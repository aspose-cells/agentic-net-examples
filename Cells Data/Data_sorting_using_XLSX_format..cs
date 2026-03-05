using System;
using Aspose.Cells;

namespace AsposeCellsSortingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Add headers and sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Pen");
            cells["B2"].PutValue("Stationery");
            cells["C2"].PutValue(1.20);

            cells["A3"].PutValue("Notebook");
            cells["B3"].PutValue("Stationery");
            cells["C3"].PutValue(2.50);

            cells["A4"].PutValue("Apple");
            cells["B4"].PutValue("Food");
            cells["C4"].PutValue(0.80);

            cells["A5"].PutValue("Bread");
            cells["B5"].PutValue("Food");
            cells["C5"].PutValue(1.00);

            // Configure the DataSorter
            DataSorter sorter = wb.DataSorter;
            sorter.HasHeaders = true;               // First row contains headers
            sorter.Key1 = 2;                         // Sort by the third column (Price)
            sorter.Order1 = SortOrder.Ascending;    // Ascending order

            // Define the area to sort (including headers)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 4,
                EndColumn = 2
            };

            // Perform the sort operation
            sorter.Sort(cells, area);

            // Save the sorted workbook in XLSX format
            wb.Save("SortedProducts.xlsx");
        }
    }
}