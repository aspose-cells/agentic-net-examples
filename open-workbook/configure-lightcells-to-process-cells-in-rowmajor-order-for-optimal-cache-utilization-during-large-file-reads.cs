// Title: Row‑Major LightCells Processing with FileCache in Aspose.Cells (C#)
// Description: Shows how to set up a custom LightCellsDataHandler that iterates worksheets, rows, and cells in natural row‑major order, releases each cell from memory, and applies the FileCache memory setting for optimal cache usage when reading large Excel workbooks.
// Keywords: Aspose.Cells LightCells C# | row major processing | FileCache memory setting | large Excel file streaming | custom LightCellsDataHandler | low‑memory Excel read | .NET Excel memory optimization | process cells sequentially | Aspose.Cells performance tuning
// Common Searches: Aspose.Cells LightCells row major order example | How to use MemorySetting.FileCache with LightCells | Release cell memory during large workbook read Aspose | C# stream large Excel file with LightCells | Custom LightCellsDataHandler tutorial
// Developer Intent: Read a massive workbook efficiently by processing cells row‑by‑row with LightCells, freeing each cell after use, and leveraging FileCache for better cache performance.
// Use Cases: Log or transform every cell in a multi‑gigabyte XLSX file while keeping RAM usage minimal. | Generate row‑wise reports or extracts from huge spreadsheets without loading the entire workbook into memory. | Convert or migrate large Excel workbooks to another format using a streaming approach that maximizes cache efficiency.
// AI Prompts: Give a C# snippet that prints each cell’s address and value in row‑major order using Aspose.Cells LightCells and frees memory after processing. | Explain how to modify the RowMajorLightCellsHandler to skip empty cells while still using MemorySetting.FileCache. | Show how to add progress‑percentage reporting to the LightCells row‑major handler for processing very large workbooks.

using System;
using Aspose.Cells;

namespace LightCellsRowMajorExample
{
    // Custom handler that processes cells in row‑major order.
    // It simply iterates through each sheet, row and cell sequentially.
    // ProcessCell returns false so cells are not kept in memory after processing,
    // which together with FileCache memory setting gives optimal cache utilization.
    // Shows how to set up a custom LightCellsDataHandler that iterates worksheets, rows, and cells in natural row‑major order, releases each cell from memory, and applies the FileCache memory setting for optimal cache usage when reading large Excel workbooks.
    public class RowMajorLightCellsHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet)
        {
            // Process all worksheets.
            return true;
        }

        public bool StartRow(int rowIndex)
        {
            // Process rows in natural order.
            return true;
        }

        public bool ProcessRow(Row row)
        {
            // No special row processing needed.
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            // Process all cells in the current row.
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            // Example: you could read or transform the cell here.
            // Returning false releases the cell from the in‑memory model.
            // This keeps memory usage low for large files.
            // For demonstration we just output the address and value.
            Console.WriteLine($"Processing {cell.Name}: {cell.Value}");
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the large source workbook.
            string sourcePath = "LargeFile_original.xlsx";
            // Path where the processed workbook will be saved.
            string outputPath = "ProcessedLargeFile.xlsx";

            // Create load options and assign the custom LightCellsDataHandler.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new RowMajorLightCellsHandler();

            // Load the workbook in LightCells mode.
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Configure each worksheet to use FileCache memory setting for better cache usage.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells.MemorySetting = MemorySetting.FileCache;
            }

            // Save the processed workbook.
            workbook.Save(outputPath);
        }
    }
}
