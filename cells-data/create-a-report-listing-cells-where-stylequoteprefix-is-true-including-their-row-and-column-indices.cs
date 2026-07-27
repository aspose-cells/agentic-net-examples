using System;
using Aspose.Cells;

class QuotePrefixReport
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet where data is stored
        Worksheet dataSheet = workbook.Worksheets[0];
        Cells dataCells = dataSheet.Cells;

        // Add a new worksheet to hold the report
        int reportSheetIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
        reportSheet.Name = "QuotePrefixReport";

        // Write header row in the report sheet
        reportSheet.Cells[0, 0].PutValue("Cell Address");
        reportSheet.Cells[0, 1].PutValue("Row Index");
        reportSheet.Cells[0, 2].PutValue("Column Index");

        int reportRow = 1; // Start writing data from the second row

        // Determine the used range of the data sheet
        int maxRow = dataCells.MaxDataRow;
        int maxCol = dataCells.MaxDataColumn;

        // Iterate through each cell in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = dataCells[row, col];
                if (cell != null && cell.GetStyle().QuotePrefix)
                {
                    // Convert row/column indices to Excel address (e.g., "B10")
                    string address = CellsHelper.CellIndexToName(row, col);

                    // Record the cell information in the report sheet
                    reportSheet.Cells[reportRow, 0].PutValue(address);
                    reportSheet.Cells[reportRow, 1].PutValue(row);
                    reportSheet.Cells[reportRow, 2].PutValue(col);
                    reportRow++;
                }
            }
        }

        // Save the workbook with the added report worksheet
        workbook.Save("output_with_report.xlsx");
    }
}