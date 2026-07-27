using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();               // create
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ------------------------------------------------------------
        // Sample data – this section can be removed when using a real file
        // ------------------------------------------------------------
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue("Text");
        cells["B1"].PutValue(5);
        cells["B2"].PutValue(15);
        cells["C1"].PutValue(100);
        cells["C2"].PutValue(-50);
        cells["C3"].PutValue(30);
        // ------------------------------------------------------------

        // Dictionary to store the maximum numeric value per column (key = column index)
        Dictionary<int, double> maxValuesPerColumn = new Dictionary<int, double>();

        // Iterate through each column using the Columns enumerator
        foreach (Column column in worksheet.Cells.Columns)
        {
            int colIndex = column.Index; // Column index (0‑based)

            // Find the last used cell in this column
            Cell lastCell = worksheet.Cells.EndCellInColumn(colIndex);
            int lastRow = lastCell.Row;

            double columnMax = double.MinValue;
            bool hasNumeric = false;

            // Scan all rows up to the last used row in this column
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = worksheet.Cells[row, colIndex];

                // Check if the cell contains a numeric value
                if (cell.Value != null && cell.Type == CellValueType.IsNumeric)
                {
                    double val = Convert.ToDouble(cell.Value);
                    if (!hasNumeric || val > columnMax)
                    {
                        columnMax = val;
                        hasNumeric = true;
                    }
                }
            }

            // Record the maximum if at least one numeric cell was found
            if (hasNumeric)
            {
                maxValuesPerColumn[colIndex] = columnMax;
            }
        }

        // Output the results
        foreach (var kvp in maxValuesPerColumn)
        {
            Console.WriteLine($"Column {kvp.Key} – Max Numeric Value: {kvp.Value}");
        }

        // Save the workbook (optional, demonstrates the required save rule)
        workbook.Save("ColumnsMaxValues.xlsx");           // save
    }
}