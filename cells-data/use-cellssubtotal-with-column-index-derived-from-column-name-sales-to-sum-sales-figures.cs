using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data with a header row that includes "Sales"
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    cells[i + 1, j].PutValue(data[i, j]);
                }
            }

            // Determine the zero‑based column index of the header named "Sales"
            int salesColumnIndex = -1;
            int headerRow = 0;
            int lastColumn = cells.MaxColumn;
            for (int col = 0; col <= lastColumn; col++)
            {
                if (cells[headerRow, col].StringValue.Equals("Sales", StringComparison.OrdinalIgnoreCase))
                {
                    salesColumnIndex = col;
                    break;
                }
            }

            if (salesColumnIndex == -1)
            {
                Console.WriteLine("Column \"Sales\" not found.");
                return;
            }

            // Define the range that contains the data (including header)
            int startRow = 0;
            int endRow = cells.MaxDataRow; // last row with data
            int startColumn = 0;
            int endColumn = cells.MaxDataColumn; // last column with data

            CellArea area = new CellArea
            {
                StartRow = startRow,
                StartColumn = startColumn,
                EndRow = endRow,
                EndColumn = endColumn
            };

            // Apply subtotal:
            // - Group by the first column (Region) -> index 0
            // - Use SUM function
            // - Add subtotal for the "Sales" column discovered above
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { salesColumnIndex });

            // Save the workbook
            workbook.Save("SubtotalBySalesColumn.xlsx");
        }
    }
}