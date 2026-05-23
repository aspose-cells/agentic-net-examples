using System;
using Aspose.Cells;

class ExportWorkbookToCsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (replace with your actual data source)
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue("Product A");
        cells["B2"].PutValue(1234.5678);   // Example numeric value
        cells["A3"].PutValue("Product B");
        cells["B3"].PutValue(987.6543);    // Example numeric value

        // Round all numeric cells to two decimal places
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    double roundedValue = Math.Round(cell.DoubleValue, 2);
                    cell.PutValue(roundedValue);
                }
            }
        }

        // Save the workbook as CSV for financial reporting
        workbook.Save("FinancialReport.csv", SaveFormat.Csv);
    }
}