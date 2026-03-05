using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ConvertTableToRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data (5 columns, 5 data rows)
        for (int col = 0; col < 5; col++)
        {
            sheet.Cells[0, col].PutValue($"Column {col + 1}"); // header
            for (int row = 1; row <= 5; row++)
            {
                sheet.Cells[row, col].PutValue(row * (col + 1));
            }
        }

        // Add a table (ListObject) that covers the populated range
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int tableIndex = sheet.ListObjects.Add(0, 0, 5, 4, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Enable totals row and set a totals calculation for the first column
        table.ShowTotals = true;
        table.ListColumns[0].TotalsCalculation = TotalsCalculation.Sum;

        // Create conversion options – specify the last row index to be converted
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 4 // Convert rows 0 through 4 (5 rows total)
        };

        // Convert the table to a normal range using the options
        table.ConvertToRange(options);

        // Save the modified workbook to an XLSX file
        workbook.Save("ConvertedTableToRange.xlsx");
    }
}