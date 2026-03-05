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

            // Add header row
            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Quantity");
            ws.Cells["C1"].PutValue("Price");

            // Add sample data rows
            ws.Cells["A2"].PutValue("Apple");
            ws.Cells["B2"].PutValue(50);
            ws.Cells["C2"].PutValue(1.2);

            ws.Cells["A3"].PutValue("Banana");
            ws.Cells["B3"].PutValue(30);
            ws.Cells["C3"].PutValue(0.8);

            ws.Cells["A4"].PutValue("Orange");
            ws.Cells["B4"].PutValue(40);
            ws.Cells["C4"].PutValue(1.0);

            // Configure the DataSorter
            DataSorter sorter = wb.DataSorter;
            sorter.HasHeaders = true;               // First row contains column names
            sorter.Key1 = 2;                         // First sort key: column C (Price)
            sorter.Order1 = SortOrder.Descending;   // Sort Price descending
            sorter.Key2 = 1;                         // Second sort key: column B (Quantity)
            sorter.Order2 = SortOrder.Ascending;    // Sort Quantity ascending

            // Define the area to be sorted (including header row)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 3,
                EndColumn = 2
            };

            // Perform the sort operation
            sorter.Sort(ws.Cells, area);

            // Save the sorted workbook in XLSX format
            wb.Save("SortedProducts.xlsx", SaveFormat.Xlsx);
        }
    }
}