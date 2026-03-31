using System;
using Aspose.Cells;

class LargeExcelWriter
{
    static void Main()
    {
        // Instantiate an empty workbook (default Xlsx format)
        Workbook workbook = new Workbook();

        // Create a custom LightCellsDataProvider that streams data row by row
        var dataProvider = new MyLightCellsDataProvider();

        // Configure save options to use LightCells mode with the custom provider
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = dataProvider
        };

        // Save the workbook using the LightCells API (efficient for large files)
        workbook.Save("LargeFile.xlsx", saveOptions);
    }

    // Custom implementation of LightCellsDataProvider
    class MyLightCellsDataProvider : LightCellsDataProvider
    {
        private int currentRow = -1;
        private int currentCol = -1;
        private const int totalRows = 100_000; // Example large row count
        private const int totalCols = 10;      // Example column count per row

        // Process only the first worksheet (index 0)
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Return the next row index to be saved, or -1 when done
        public int NextRow()
        {
            currentRow++;
            currentCol = -1; // Reset column for new row
            return currentRow < totalRows ? currentRow : -1;
        }

        // Called when a new row starts; no special handling needed here
        public void StartRow(Row row) { }

        // Return the next column index within the current row, or -1 when done
        public int NextCell()
        {
            currentCol++;
            return currentCol < totalCols ? currentCol : -1;
        }

        // Populate the cell with data; here we write a simple identifier
        public void StartCell(Cell cell)
        {
            cell.PutValue($"R{currentRow + 1}C{currentCol + 1}");
        }

        // Indicate whether strings should be gathered into a global pool
        public bool IsGatherString()
        {
            return true;
        }
    }
}