using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsFormulaExample
{
    // Custom LightCellsDataProvider that streams data and formulas
    public class FormulaDataProvider : LightCellsDataProvider
    {
        private int _currentRow = -1;
        private int _currentCol = -1;
        private const int MaxRows = 5; // Header + 4 data rows
        private const int MaxCols = 3; // ID, Value, Formula

        // Process only the first worksheet (index 0)
        public bool StartSheet(int sheetIndex) => sheetIndex == 0;

        // Return next row index or -1 when finished
        public int NextRow()
        {
            _currentRow++;
            _currentCol = -1; // reset column for new row
            return _currentRow < MaxRows ? _currentRow : -1;
        }

        public void StartRow(Row row) { }

        // Return next column index or -1 when finished
        public int NextCell()
        {
            _currentCol++;
            return _currentCol < MaxCols ? _currentCol : -1;
        }

        // Fill cell with value or formula
        public void StartCell(Cell cell)
        {
            // Header row
            if (_currentRow == 0)
            {
                switch (_currentCol)
                {
                    case 0: cell.PutValue("ID"); break;
                    case 1: cell.PutValue("Value"); break;
                    case 2: cell.PutValue("Formula"); break;
                }
                return;
            }

            // Data rows
            switch (_currentCol)
            {
                case 0: // ID column
                    cell.PutValue(_currentRow); // IDs start from 1
                    break;
                case 1: // Value column (simple numeric value)
                    cell.PutValue(_currentRow * 10); // e.g., 10,20,30...
                    break;
                case 2: // Formula column: double the value in column B of the same row
                    // Set the formula; value will be calculated later
                    string formula = $"=B{_currentRow + 1}*2";
                    cell.Formula = formula;
                    break;
            }
        }

        // Gather strings into the global pool (optional, true improves performance for many strings)
        public bool IsGatherString() => true;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create an empty workbook (default worksheet is present)
                Workbook wb = new Workbook();

                // 2. Configure save options to use LightCells mode with our custom provider
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
                {
                    LightCellsDataProvider = new FormulaDataProvider()
                };

                // 3. Save the workbook to a memory stream using LightCells streaming
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.Save(stream, saveOptions);

                    // Reset stream position for reading
                    stream.Position = 0;

                    // 4. Load the saved workbook normally (full model) to verify formulas
                    Workbook loadedWb = new Workbook(stream);

                    // 5. Calculate all formulas
                    loadedWb.CalculateFormula();

                    // 6. Output the resulting values to the console
                    Worksheet ws = loadedWb.Worksheets[0];
                    Cells cells = ws.Cells;

                    Console.WriteLine("ID\tValue\tFormulaResult");
                    for (int row = 1; row < 5; row++) // rows 1..4 contain data (row 0 is header)
                    {
                        int id = cells[row, 0].IntValue;
                        double val = cells[row, 1].DoubleValue;
                        double formulaResult = cells[row, 2].DoubleValue;
                        Console.WriteLine($"{id}\t{val}\t{formulaResult}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}