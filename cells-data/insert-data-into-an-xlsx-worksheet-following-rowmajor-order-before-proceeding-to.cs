using System;
using Aspose.Cells;

namespace AsposeCellsRowMajorThenColumnWise
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Insert data in row‑major order (horizontal import)
            // ------------------------------------------------------------
            // Sample data: header row followed by two data rows
            // The array will be placed horizontally starting at cell A1
            object[] rowMajorData = new object[]
            {
                "ID", "Name", "Price",          // Header
                1,   "Product A", 19.99,        // Row 2
                2,   "Product B", 29.99         // Row 3
            };
            // false => import horizontally (row‑major)
            cells.ImportObjectArray(rowMajorData, firstRow: 0, firstColumn: 0, isVertical: false);

            // ------------------------------------------------------------
            // 2. Insert data in column‑wise order (vertical import)
            // ------------------------------------------------------------
            // Sample data: a simple list that will be placed vertically
            // Starting at cell A6 (row index 5, column index 0)
            object[] columnWiseData = new object[]
            {
                "Category",
                "Electronics",
                "Furniture",
                "Stationery"
            };
            // true => import vertically (column‑wise)
            cells.ImportObjectArray(columnWiseData, firstRow: 5, firstColumn: 0, isVertical: true);

            // Save the workbook to an XLSX file
            workbook.Save("RowMajorThenColumnWise.xlsx");
        }
    }
}