using Aspose.Cells;
using System;

class SubtotalByColumnName
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with a header row that includes the "Sales" column
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

        for (int r = 0; r < data.GetLength(0); r++)
        {
            for (int c = 0; c < data.GetLength(1); c++)
            {
                cells[r + 1, c].PutValue(data[r, c]);
            }
        }

        // Determine the zero‑based column index of the header named "Sales"
        int salesColumnIndex = -1;
        int headerRow = 0;
        int lastColumn = cells.MaxDataColumn;

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
            Console.WriteLine("Column 'Sales' not found.");
            return;
        }

        // Define the cell area that contains the whole data set (including header)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = cells.MaxDataRow,
            EndColumn = cells.MaxDataColumn
        };

        // Group by the first column (Region) and sum the values in the "Sales" column
        int groupByColumnIndex = 0; // Region column
        cells.Subtotal(area, groupByColumnIndex, ConsolidationFunction.Sum, new int[] { salesColumnIndex });

        // Save the workbook
        workbook.Save("SubtotalBySalesColumn.xlsx");
    }
}