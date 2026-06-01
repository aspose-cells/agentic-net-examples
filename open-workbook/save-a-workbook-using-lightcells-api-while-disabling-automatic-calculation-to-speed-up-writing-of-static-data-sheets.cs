using System;
using Aspose.Cells;

namespace LightCellsSaveExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Disable automatic formula calculation on save and on open
            workbook.Settings.FormulaSettings.CalculateOnSave = false;
            workbook.Settings.FormulaSettings.CalculateOnOpen = false;

            // Configure LightCells save options with a custom data provider
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new CustomLightCellsDataProvider()
            };

            // Save the workbook using LightCells API
            workbook.Save("output_lightcells.xlsx", saveOptions);
        }
    }

    // Custom LightCellsDataProvider that writes a simple 5x5 numeric grid
    public class CustomLightCellsDataProvider : LightCellsDataProvider
    {
        private const int MaxRows = 5;
        private const int MaxCols = 5;
        private int currentRow = -1;
        private int currentCol = -1;

        public bool StartSheet(int sheetIndex)
        {
            // Process only the first worksheet
            return sheetIndex == 0;
        }

        public int NextRow()
        {
            if (currentRow < MaxRows - 1)
            {
                currentRow++;
                currentCol = -1; // reset column for new row
                return currentRow;
            }
            return -1; // no more rows
        }

        public void StartRow(Row row)
        {
            // No special row handling needed
        }

        public int NextCell()
        {
            if (currentCol < MaxCols - 1)
            {
                currentCol++;
                return currentCol;
            }
            return -1; // no more cells in this row
        }

        public void StartCell(Cell cell)
        {
            // Populate cell with a simple numeric value
            cell.PutValue(currentRow * 10 + currentCol);
        }

        public bool IsGatherString()
        {
            // No string gathering needed for this example
            return false;
        }
    }
}