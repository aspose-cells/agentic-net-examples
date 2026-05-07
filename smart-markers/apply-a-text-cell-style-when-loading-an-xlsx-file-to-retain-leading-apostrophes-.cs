using Aspose.Cells;
using System;

class RetainApostropheDemo
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        Workbook workbook = new Workbook(inputPath);
        workbook.Settings.QuotePrefixToStyle = true;

        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // Skip blank cells
                if (cell.Type == CellValueType.IsNull)
                    continue;

                string str = cell.StringValue;
                if (!string.IsNullOrEmpty(str) && str.StartsWith("'"))
                {
                    Style style = cell.GetStyle();
                    style.QuotePrefix = true;
                    cell.SetStyle(style);
                }
            }
        }

        workbook.Save("output.xlsx");
    }
}