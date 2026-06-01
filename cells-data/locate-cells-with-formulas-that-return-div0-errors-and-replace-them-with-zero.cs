using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure all formulas are calculated so that error values are available
        workbook.CalculateFormula();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxRow;
            int maxColumn = cells.MaxColumn;

            // Scan every cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a formula that resulted in a #DIV/0! error
                    if (cell.IsErrorValue && cell.StringValue == "#DIV/0!")
                    {
                        // Replace the error with zero
                        cell.PutValue(0);
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}