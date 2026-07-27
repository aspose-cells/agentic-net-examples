// Title: Save a Workbook with LightCells API and Disable Formula Calculation (C# Aspose.Cells)
// Description: Shows how to create a workbook, turn off automatic formula calculation, implement a LightCellsDataProvider for static data, set OoxmlSaveOptions to LightCells mode, and save the file as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | LightCells API | C# | disable CalculateOnSave | OoxmlSaveOptions | static data export | workbook save performance | formula calculation off | LightCellsDataProvider example | .NET Excel generation
// Common Searches: Aspose.Cells LightCells save static data C# | disable formula calculation on save Aspose.Cells | LightCellsDataProvider example .NET | speed up workbook save with LightCells | save workbook without recalculation Aspose.Cells
// Developer Intent: The developer wants to write static data to an XLSX file with LightCells while preventing formula recalculation to accelerate the save operation.
// Use Cases: Export large static lookup tables without triggering formulas | Generate pre‑filled reports where calculations are unnecessary | Create template workbooks for downstream processing | Batch‑write millions of rows with minimal overhead
// AI Prompts: Provide a C# snippet that saves a workbook using LightCells and sets CalculateOnSave to false for faster static data writing. | Explain how to extend SimpleStaticDataProvider to write to multiple worksheets while keeping automatic calculation disabled. | Show how to apply custom cell formatting during LightCells saving without enabling formula calculation. | Describe how to combine LightCells export with CSV conversion in Aspose.Cells while controlling calculation settings.

using System;
using Aspose.Cells;

// Shows how to create a workbook, turn off automatic formula calculation, implement a LightCellsDataProvider for static data, set OoxmlSaveOptions to LightCells mode, and save the file as XLSX using Aspose.Cells for .NET.
class LightCellsSaveDemo
{
    static void Main()
    {
        // Create a new empty workbook
        Workbook workbook = new Workbook();

        // Disable automatic formula calculation on save (speed up static data saving)
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // Create a LightCells data provider that supplies static data
        LightCellsDataProvider provider = new SimpleStaticDataProvider();

        // Configure OoxmlSaveOptions to use LightCells mode
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = provider
        };

        // Save the workbook using LightCells API
        workbook.Save("StaticDataLightCells.xlsx", saveOptions);
    }

    // Simple implementation of LightCellsDataProvider that writes a small static table
    private class SimpleStaticDataProvider : LightCellsDataProvider
    {
        private int currentRow = -1;
        private int currentCol = -1;

        // Sample static data to be written
        private readonly string[,] data = new string[,]
        {
            { "ID", "Name", "Qty" },
            { "1", "Item A", "10" },
            { "2", "Item B", "20" },
            { "3", "Item C", "15" }
        };

        // Process only the first worksheet
        public bool StartSheet(int sheetIndex) => sheetIndex == 0;

        // Return the next row index or -1 when done
        public int NextRow()
        {
            currentRow++;
            currentCol = -1;
            return currentRow < data.GetLength(0) ? currentRow : -1;
        }

        // No special row handling required
        public void StartRow(Row row) { }

        // Return the next column index or -1 when done
        public int NextCell()
        {
            currentCol++;
            return currentCol < data.GetLength(1) ? currentCol : -1;
        }

        // Set the cell value from the static data array
        public void StartCell(Cell cell)
        {
            cell.PutValue(data[currentRow, currentCol]);
        }

        // Strings are not gathered into a global pool for this simple example
        public bool IsGatherString() => false;
    }
}
