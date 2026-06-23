using System;
using Aspose.Cells;

class SubtotalExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells N10:O10 (zero‑based indices: row 9, column 13)
        // totalRows = 1 (single row), totalColumns = 2 (N and O)
        cells.Merge(9, 13, 1, 2);

        // Put the label "Subtotal" in the merged cell (upper‑left corner of the range)
        cells[9, 13].PutValue("Subtotal");

        // Determine the last row that contains data in column N (index 13)
        int lastDataRow = cells.GetLastDataRow(13);
        // Ensure the range starts after the label row (row 11 in Excel = index 10)
        if (lastDataRow < 10) lastDataRow = 10;

        // Build a SUM formula for column N from row 11 to the last data row
        // Excel rows are 1‑based, so add 1 to the zero‑based indices
        string sumFormula = $"=SUM(N11:N{lastDataRow + 1})";

        // Assign the formula to the merged cell
        cells[9, 13].Formula = sumFormula;

        // Calculate the formula so the result is stored in the cell
        workbook.CalculateFormula();

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}