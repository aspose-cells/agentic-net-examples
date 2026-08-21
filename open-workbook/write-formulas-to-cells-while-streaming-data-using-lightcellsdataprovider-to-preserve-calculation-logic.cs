// Title: Write Excel formulas while streaming rows with LightCellsDataProvider in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, configure OoxmlSaveOptions for LightCells mode, implement a custom LightCellsDataProvider that streams headers, data rows, and a formula (Price * Qty) using SetFormula, save the file, reload it, recalculate formulas, and output the results.
// Keywords: Aspose.Cells | .NET | LightCells | LightCellsDataProvider | streaming data | Excel formula | SetFormula | memory‑efficient export | OoxmlSaveOptions | large worksheet generation
// Common Searches: Aspose.Cells LightCellsDataProvider write formula | set formula while streaming cells Aspose.Cells .NET | memory efficient Excel export with formulas | custom LightCellsDataProvider example | preserve formulas after saving with LightCells mode
// Developer Intent: Generate an Excel file using LightCells mode, stream cell values and formulas row‑by‑row, and ensure the formulas remain active after the workbook is saved.
// Use Cases: Export massive data sets with calculated columns without loading the entire sheet into memory. | Create invoice or sales reports where each row includes a total formula while streaming rows for performance. | Build lightweight Excel files that define formulas on the fly for dynamic calculations.
// AI Prompts: Show how to extend the FormulaDataProvider to add a Discount column that uses a formula referencing Price and Qty. | Provide a LightCellsDataProvider example that writes conditional‑formatting formulas while streaming data. | Explain how to retrieve, modify, and evaluate formulas after loading a workbook saved in LightCells mode.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, configure OoxmlSaveOptions for LightCells mode, implement a custom LightCellsDataProvider that streams headers, data rows, and a formula (Price * Qty) using SetFormula, save the file, reload it, recalculate formulas, and output the results.
class Program
{
    static void Main()
    {
        // Create an empty workbook
        Workbook wb = new Workbook();

        // Configure save options to use LightCells mode with a custom provider
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = new FormulaDataProvider()
        };

        // Save the workbook (the provider streams cell data and writes formulas)
        string filePath = "LightCellsFormulas.xlsx";
        wb.Save(filePath, saveOptions);

        // Load the saved workbook normally to verify that formulas were preserved
        Workbook loaded = new Workbook(filePath);
        loaded.CalculateFormula(); // Evaluate all formulas

        // Output the calculated results
        var cells = loaded.Worksheets[0].Cells;
        for (int r = 1; r <= 5; r++)
        {
            Console.WriteLine($"Row {r}: ID={cells[r, 0].IntValue}, Price={cells[r, 1].DoubleValue}, Qty={cells[r, 2].IntValue}, Total={cells[r, 3].DoubleValue}");
        }
    }

    // Custom LightCellsDataProvider that streams data and writes formulas
    class FormulaDataProvider : LightCellsDataProvider
    {
        private int currentRow = -1;
        private int currentCol = -1;
        private const int totalRows = 6; // 1 header row + 5 data rows
        private const int totalCols = 4; // ID, Price, Qty, Total

        public bool StartSheet(int sheetIndex)
        {
            // Process only the first worksheet
            return sheetIndex == 0;
        }

        public int NextRow()
        {
            currentRow++;
            currentCol = -1;
            return currentRow < totalRows ? currentRow : -1;
        }

        public void StartRow(Row row)
        {
            // No special row handling required
        }

        public int NextCell()
        {
            currentCol++;
            return currentCol < totalCols ? currentCol : -1;
        }

        public void StartCell(Cell cell)
        {
            // Header row
            if (currentRow == 0)
            {
                switch (currentCol)
                {
                    case 0: cell.PutValue("ID"); break;
                    case 1: cell.PutValue("Price"); break;
                    case 2: cell.PutValue("Qty"); break;
                    case 3: cell.PutValue("Total"); break;
                }
                return;
            }

            // Data rows (rows 1..5)
            int dataRow = currentRow; // 1‑based for Excel formulas
            switch (currentCol)
            {
                case 0:
                    // ID column
                    cell.PutValue(dataRow);
                    break;
                case 1:
                    // Price column (example values)
                    cell.PutValue(10 + dataRow * 2);
                    break;
                case 2:
                    // Qty column (example values)
                    cell.PutValue(dataRow % 3 + 1);
                    break;
                case 3:
                    // Total column – set a formula that multiplies Price * Qty
                    string formula = $"=B{dataRow + 1}*C{dataRow + 1}";
                    // Use SetFormula to store the formula; value will be calculated later
                    cell.SetFormula(formula, new FormulaParseOptions(), null);
                    break;
            }
        }

        public bool IsGatherString()
        {
            // Gather strings into the global string pool for efficiency
            return true;
        }
    }
}
