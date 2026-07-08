using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data in columns J (index 9) and K (index 10)
            // Header row
            worksheet.Cells["J1"].PutValue("Category");
            worksheet.Cells["K1"].PutValue("Value");

            // Sample data rows
            object[,] data = new object[,]
            {
                { "A", 10 },
                { "A", 20 },
                { "B", 15 },
                { "B", 25 },
                { "C", 30 }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                // Row index is i+1 because row 0 contains headers
                worksheet.Cells[i + 1, 9].PutValue(data[i, 0]); // Column J
                worksheet.Cells[i + 1, 10].PutValue(data[i, 1]); // Column K
            }

            // Define the cell area that includes the data (J1:K6)
            // StartRow = 0, StartColumn = 9 (J), EndRow = 5, EndColumn = 10 (K)
            CellArea area = CellArea.CreateCellArea(0, 9, 5, 10);

            // Add subtotals:
            // - Group by the first column of the area (Category, offset 0)
            // - Use CountNums function on the second column (Value, offset 1)
            // - No outline levels are created because we use the 4‑parameter overload
            worksheet.Cells.Subtotal(
                area,
                0,                                 // groupBy: first column in the area
                ConsolidationFunction.CountNums,   // function: CountNumbers
                new int[] { 1 }                    // totalList: apply to second column in the area
            );

            // Save the workbook
            workbook.Save("Subtotal_CountNums_ColumnK.xlsx");
        }
    }
}