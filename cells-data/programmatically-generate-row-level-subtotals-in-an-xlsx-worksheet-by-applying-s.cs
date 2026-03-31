using Aspose.Cells;
using System;

class RowLevelSubtotalDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add header row
        cells["A1"].PutValue("Department");
        cells["B1"].PutValue("Employee");
        cells["C1"].PutValue("Salary");

        // Populate sample data
        object[,] data = new object[,]
        {
            { "HR",      "Alice",   5000 },
            { "HR",      "Bob",     5500 },
            { "IT",      "Charlie", 7000 },
            { "IT",      "David",   7200 },
            { "IT",      "Eve",     6800 },
            { "Finance", "Frank",   6000 },
            { "Finance", "Grace",   6200 }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]); // Department
            cells[i + 1, 1].PutValue(data[i, 1]); // Employee
            cells[i + 1, 2].PutValue(data[i, 2]); // Salary
        }

        // Define the cell area that includes the header and all data rows
        // StartRow = 0, StartColumn = 0, EndRow = data rows count, EndColumn = 2
        CellArea area = CellArea.CreateCellArea(0, 0, data.GetLength(0), 2);

        // Apply subtotals:
        // - Group by column 0 (Department)
        // - Use SUM function
        // - Add subtotal to column 2 (Salary)
        // - Replace existing subtotals, add page breaks, place summary below data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, true, true);

        // Save the workbook with the generated subtotals
        workbook.Save("RowLevelSubtotals.xlsx");
    }
}