// Title: Write formulas while streaming rows with a custom LightCellsDataProvider in Aspose.Cells for .NET
// Description: Demonstrates how to create an empty Workbook, configure OoxmlSaveOptions for LightCells mode, and use a custom LightCellsDataProvider to stream a 6‑row, 5‑column table. The provider writes header text, populates ID, Name, Price, and Qty cells, and inserts a formula "=C{row}*D{row}" in the Total column, preserving calculation logic without loading the entire sheet into memory.
// Keywords: Aspose.Cells | LightCells | DataProvider | formula streaming | .NET | C# | memory‑efficient Excel export | stream rows to XLSX | custom LightCellsDataProvider | write formulas with LightCells | OoxmlSaveOptions
// Common Searches: Aspose.Cells LightCells write formulas | stream Excel rows with formulas .NET | custom LightCellsDataProvider example | save workbook with formulas using LightCells | memory‑efficient Excel export Aspose
// Developer Intent: Stream rows and formulas to an Excel file using a custom LightCellsDataProvider.
// Use Cases: Generate a massive sales report where each row’s Total column is a live formula, avoiding full‑sheet loading. | Export database query results to XLSX while keeping Price × Quantity calculations as formulas for downstream users. | Create a templated workbook on the fly with headers and computed columns, then distribute it for further analysis.
// AI Prompts: Provide a LightCellsDataProvider that streams millions of rows with a per‑row SUM formula while staying under a low memory footprint. | Show how to apply number formats and cell styles to formula cells streamed via LightCellsDataProvider in Aspose.Cells. | Explain how to enable automatic recalculation after saving a workbook that contains formulas written through LightCellsDataProvider.

using System;
using Aspose.Cells;

// Demonstrates how to create an empty Workbook, configure OoxmlSaveOptions for LightCells mode, and use a custom LightCellsDataProvider to stream a 6‑row, 5‑column table. The provider writes header text, populates ID, Name, Price, and Qty cells, and inserts a formula "=C{row}*D{row}" in the Total column, preserving calculation logic without loading the entire sheet into memory.
class Program
{
    static void Main()
    {
        try
        {
            // Create an empty workbook; data will be supplied by the LightCellsDataProvider
            Workbook workbook = new Workbook();

            // Configure save options to use LightCells mode with the custom provider
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new FormulaStreamingProvider()
            };

            // Save the workbook – the provider streams rows/cells with formulas into the file
            workbook.Save("FormulasStreaming.xlsx", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Custom LightCellsDataProvider that streams a small table containing formulas
    class FormulaStreamingProvider : LightCellsDataProvider
    {
        private int currentRow = -1;
        private int currentCol = -1;

        // Define the size of the table (header + 5 data rows, 5 columns)
        private const int TotalRows = 6; // 0 = header, 1‑5 = data
        private const int TotalCols = 5; // ID, Name, Price, Qty, Total

        // Process only the first worksheet (index 0)
        public bool StartSheet(int sheetIndex) => sheetIndex == 0;

        // Return the next row index to be saved; -1 signals completion
        public int NextRow()
        {
            currentRow++;
            currentCol = -1; // reset column for the new row
            return currentRow < TotalRows ? currentRow : -1;
        }

        // No special row initialization required
        public void StartRow(Row row) { }

        // Return the next column index for the current row; -1 signals end of columns
        public int NextCell()
        {
            currentCol++;
            return currentCol < TotalCols ? currentCol : -1;
        }

        // Fill the cell with either a header, a value, or a formula
        public void StartCell(Cell cell)
        {
            // Header row (row 0)
            if (currentRow == 0)
            {
                switch (currentCol)
                {
                    case 0: cell.PutValue("ID"); break;
                    case 1: cell.PutValue("Name"); break;
                    case 2: cell.PutValue("Price"); break;
                    case 3: cell.PutValue("Qty"); break;
                    case 4: cell.PutValue("Total"); break;
                }
                return;
            }

            // Data rows (rows 1‑5)
            int excelRow = currentRow + 1; // Excel rows are 1‑based
            switch (currentCol)
            {
                case 0:
                    // ID
                    cell.PutValue(currentRow);
                    break;
                case 1:
                    // Name
                    cell.PutValue($"Item {currentRow}");
                    break;
                case 2:
                    // Price (simple incremental price)
                    cell.PutValue(10 + currentRow);
                    break;
                case 3:
                    // Quantity (cycle 1‑3)
                    cell.PutValue((currentRow % 3) + 1);
                    break;
                case 4:
                    // Formula: =C{row}*D{row}
                    string formula = $"=C{excelRow}*D{excelRow}";
                    // Set the formula; the calculated value will be computed later
                    cell.Formula = formula;
                    break;
            }
        }

        // Enable global string pooling for efficiency
        public bool IsGatherString() => true;
    }
}
