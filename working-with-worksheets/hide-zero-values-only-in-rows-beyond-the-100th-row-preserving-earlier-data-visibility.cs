using System;
using Aspose.Cells;

class HideZeroBeyondRow100
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (120 rows) with some zero values
        for (int i = 0; i < 120; i++)
        {
            // Example: put zero in every 10th row
            cells[i, 0].PutValue(i % 10 == 0 ? 0 : i);
        }

        // Ensure zeros are displayed globally (default is true)
        worksheet.DisplayZeros = true;

        // Create a style that hides zero values (third section of custom format is empty)
        Style hideZeroStyle = workbook.CreateStyle();
        hideZeroStyle.Custom = "0;-0;;";

        // Apply the style only to rows beyond the 100th row that contain zero
        for (int row = 100; row <= cells.MaxDataRow; row++)
        {
            Cell cell = cells[row, 0];
            if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue == 0)
            {
                cell.SetStyle(hideZeroStyle);
            }
        }

        // Save the workbook
        workbook.Save("HideZeroBeyondRow100.xlsx");
    }
}