using System;
using Aspose.Cells;

namespace LightCellsRowMajorDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the large source workbook
            string sourcePath = "LargeFile_original.xlsx";
            // Path where the processed workbook will be saved
            string destPath = "ProcessedLargeFile.xlsx";

            // Create an instance of the custom LightCellsDataHandler
            var handler = new RowMajorHandler();

            // Configure load options to use the handler in LightCells mode
            var loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = handler;

            // Load the workbook using LightCells mode (row‑major processing)
            var workbook = new Workbook(sourcePath, loadOptions);

            // Optional: use FileCache memory setting for very large files
            workbook.Settings.MemorySetting = MemorySetting.FileCache;
            workbook.Worksheets[0].Cells.MemorySetting = MemorySetting.FileCache;

            // Save the processed workbook
            workbook.Save(destPath);
        }
    }

    // Custom LightCellsDataHandler that processes cells sequentially row by row
    public class RowMajorHandler : LightCellsDataHandler
    {
        // Called for each worksheet; return true to process the sheet
        public bool StartSheet(Worksheet sheet)
        {
            // No special sheet handling needed
            return true;
        }

        // Called before a row is read; return true to process the row
        public bool StartRow(int rowIndex)
        {
            // Process rows in the order they appear (row‑major)
            return true;
        }

        // Called after a row is read; return true to continue processing its cells
        public bool ProcessRow(Row row)
        {
            // No row‑level modifications required
            return true;
        }

        // Called before a cell is read; return true to process the cell
        public bool StartCell(int columnIndex)
        {
            // Process every cell in the current row
            return true;
        }

        // Called after a cell is read; return false to discard the cell after processing
        public bool ProcessCell(Cell cell)
        {
            // Example processing: output cell address and value
            Console.WriteLine($"Processing {cell.Name}: {cell.Value}");

            // Discard the cell to keep memory usage low
            return false;
        }
    }
}