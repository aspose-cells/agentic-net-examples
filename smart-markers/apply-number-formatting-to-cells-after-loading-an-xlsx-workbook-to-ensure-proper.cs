using System;
using Aspose.Cells;

class ApplyNumberFormatting
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a style that displays numbers with two decimal places (0.00)
        Style numberStyle = workbook.CreateStyle();
        numberStyle.Number = 2; // 2 corresponds to the "0.00" format

        // Get the used range of the worksheet
        Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

        // Apply the number style only to cells that contain numeric values
        for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
        {
            for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
            {
                Cell cell = worksheet.Cells[row, col];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    cell.SetStyle(numberStyle);
                }
            }
        }

        // Save the workbook with the applied number formatting
        workbook.Save("output.xlsx");
    }
}