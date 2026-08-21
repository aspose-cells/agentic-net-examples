// Title: Read Large Excel Files Row‑by‑Row with LightCellsDataHandler in Aspose.Cells for .NET
// Description: Shows how to create a custom LightCellsDataHandler that walks worksheets, rows, and cells in row‑major order, logs each cell’s address and value, and returns false to release the cell from memory. Includes assigning the handler via LoadOptions, enabling FileCache memory setting, and saving the processed workbook for low‑memory processing of massive Excel files.
// Keywords: Aspose.Cells | LightCells | LightCellsDataHandler | row-major order | .NET | C# | large Excel file | memory optimization | FileCache | streaming Excel | low‑memory processing
// Common Searches: Aspose.Cells LightCells row major processing | How to read large Excel files with low memory using LightCells | Custom LightCellsDataHandler example C# | Discard cells after processing Aspose.Cells | Enable FileCache memory setting Aspose.Cells
// Developer Intent: Implement a LightCellsDataHandler that processes cells sequentially row‑by‑row and frees each cell from memory while loading a large workbook.
// Use Cases: Stream a multi‑gigabyte workbook, output each cell to a log, and keep RAM usage minimal. | Apply on‑the‑fly transformations (e.g., calculations or formatting) while reading a huge file, then save the result without loading the entire sheet into memory. | Generate aggregate statistics such as row totals or distinct values by processing cells row‑major and discarding them after use.
// AI Prompts: Create a LightCellsDataHandler that writes each processed cell to a CSV file and returns false to free memory. | Modify the RowMajorLightCellsHandler to skip empty cells and only process numeric values, then save the workbook. | Write a LightCellsDataHandler that counts non‑empty cells per row, stores the counts in a dictionary, and discards the cells after processing.

using System;
using Aspose.Cells;

namespace LightCellsRowMajorExample
{
    // Custom handler that processes cells in row‑major order.
    // It simply iterates through each sheet, row and cell sequentially.
    // ProcessCell returns false so cells are not kept in memory after processing,
    // which reduces memory usage for large files.
    // Shows how to create a custom LightCellsDataHandler that walks worksheets, rows, and cells in row‑major order, logs each cell’s address and value, and returns false to release the cell from memory. Includes assigning the handler via LoadOptions, enabling FileCache memory setting, and saving the processed workbook for low‑memory processing of massive Excel files.
    public class RowMajorLightCellsHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet)
        {
            // Process all worksheets.
            return true;
        }

        public bool StartRow(int rowIndex)
        {
            // Process every row in order.
            return true;
        }

        public bool ProcessRow(Row row)
        {
            // No special row processing needed; continue.
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            // Process every cell in the current row.
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            // Example: you could read or transform the cell here.
            // Returning false discards the cell from the in‑memory model,
            // keeping only the processed result.
            // For demonstration, we just output the cell address and value.
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
            string targetPath = "ProcessedLargeFile.xlsx";

            // Create load options and assign the custom LightCellsDataHandler.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new RowMajorLightCellsHandler();

            // Load the workbook in LightCells mode using the handler.
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Optional: set memory setting to FileCache for large files.
            workbook.Worksheets[0].Cells.MemorySetting = MemorySetting.FileCache;

            // Save the processed workbook.
            workbook.Save(targetPath);
        }
    }
}
