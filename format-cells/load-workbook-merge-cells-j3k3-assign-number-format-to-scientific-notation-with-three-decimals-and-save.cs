using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") to load

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells J3:K3 (zero‑based indices: row 2, column 9, 1 row, 2 columns)
        cells.Merge(firstRow: 2, firstColumn: 9, totalRows: 1, totalColumns: 2);

        // Apply scientific notation with three decimal places to the merged cell
        // J3 is the upper‑left cell of the merged range
        Cell mergedCell = cells["J3"];
        Style style = mergedCell.GetStyle();

        // Use a custom format for three decimal places in scientific notation
        style.Custom = "0.000E+00";   // e.g., 1.234E+03
        // Alternatively, you could set style.Number = 11 (default scientific) and then adjust precision,
        // but a custom format gives exact control over decimal places.

        mergedCell.SetStyle(style);

        // Optionally put a numeric value to see the formatting
        mergedCell.PutValue(12345.6789);

        // Save the workbook
        workbook.Save("MergedScientific.xlsx");
    }
}