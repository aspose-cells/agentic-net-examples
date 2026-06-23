using System;
using Aspose.Cells;

namespace LightCellsRowCopyDemo
{
    // LightCellsDataProvider that delegates all saving to the default data model.
    // Returning false from StartSheet tells Aspose.Cells to use its normal saving mechanism,
    // while the workbook is still saved in LightCells mode to reduce memory usage.
    public class PassThroughLightCellsDataProvider : LightCellsDataProvider
    {
        public bool StartSheet(int sheetIndex)
        {
            // Do not provide custom data for any sheet.
            return false;
        }

        public int NextRow()
        {
            // No rows are supplied by the provider.
            return -1;
        }

        public void StartRow(Row row)
        {
            // No custom row handling required.
        }

        public int NextCell()
        {
            // No cells are supplied by the provider.
            return -1;
        }

        public void StartCell(Cell cell)
        {
            // No custom cell handling required.
        }

        public bool IsGatherString()
        {
            // Use default string handling.
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook containing the data.
            string sourceFile = "source.xlsx";

            // Path where the new workbook with selected rows will be saved.
            string destinationFile = "selectedRows.xlsx";

            // Load the source workbook.
            Workbook sourceWorkbook = new Workbook(sourceFile);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range of rows to copy.
            // Example: copy rows 2 through 6 (zero‑based indices 1 to 5).
            int sourceStartRow = 1;          // first row to copy (index 1)
            int rowsToCopy = 5;              // number of rows to copy

            // Create a new (empty) workbook for the destination.
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Copy the selected rows from the source sheet to the destination sheet.
            // Destination start row is 0 (the first row of the new sheet).
            destinationSheet.Cells.CopyRows(
                sourceSheet.Cells,
                sourceStartRow,
                0,
                rowsToCopy);

            // Prepare save options that enable LightCells mode.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new PassThroughLightCellsDataProvider()
            };

            // Save the destination workbook using LightCells to minimize memory consumption.
            destinationWorkbook.Save(destinationFile, saveOptions);

            Console.WriteLine($"Selected rows have been copied and saved to '{destinationFile}'.");
        }
    }
}