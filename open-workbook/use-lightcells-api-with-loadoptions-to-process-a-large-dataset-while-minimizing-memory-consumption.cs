// Title: Stream a Massive Excel Workbook with Aspose.Cells LightCells API and LoadOptions (C#)
// Description: Demonstrates loading a multi‑gigabyte XLSX file in streaming mode using a custom LightCellsDataHandler, disabling KeepUnparsedData, limiting rows, logging each cell, and saving the workbook while keeping RAM usage minimal.
// Keywords: Aspose.Cells | LightCells API | LoadOptions | KeepUnparsedData false | C# streaming Excel | large workbook processing | memory‑efficient Excel read | .NET Excel handler | GitHub example | custom LightCellsDataHandler
// Common Searches: Aspose.Cells LightCells streaming example C# | How to read large Excel file without loading into memory | LoadOptions KeepUnparsedData false usage | Limit rows with LightCells StartRow method | Custom LightCellsDataHandler tutorial
// Developer Intent: Read and process a huge Excel file in a low‑memory, row‑by‑row fashion using LightCells.
// Use Cases: Extract data from a multi‑GB .xlsx without exhausting RAM. | Generate a summary of the first million rows of each sheet. | Create a trimmed copy of a large workbook after read‑only analysis. | Log every cell value to an external system while streaming.
// AI Prompts: Write a LightCellsDataHandler that copies rows matching a specific column value into a new workbook. | Show how to modify SimpleLightCellsHandler to transform cell values before saving. | Provide code to export streamed cell addresses and values to a CSV file using Aspose.Cells LightCells.

using System;
using Aspose.Cells;

namespace LightCellsProcessingDemo
{
    // Demonstrates loading a multi‑gigabyte XLSX file in streaming mode using a custom LightCellsDataHandler, disabling KeepUnparsedData, limiting rows, logging each cell, and saving the workbook while keeping RAM usage minimal.
    class Program
    {
        static void Main()
        {
            // Path to the large Excel file to be processed
            string inputPath = "LargeData.xlsx";
            // Path where the processed workbook will be saved (optional)
            string outputPath = "ProcessedLargeData.xlsx";

            // Create an instance of the custom LightCellsDataHandler
            var handler = new SimpleLightCellsHandler();

            // Configure LoadOptions to use LightCells mode and assign the handler
            var loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = handler;
            // Disable keeping unparsed data to further reduce memory consumption
            loadOptions.KeepUnparsedData = false;

            // Load the workbook in streaming (light) mode
            var workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook after processing (no modifications made in this example)
            workbook.Save(outputPath);
        }
    }

    // Custom implementation of LightCellsDataHandler for streaming processing
    public class SimpleLightCellsHandler : LightCellsDataHandler
    {
        // Called for each worksheet; return true to process the sheet
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Processing sheet: {sheet.Name}");
            return true;
        }

        // Called for each row index; return true to read the row
        public bool StartRow(int rowIndex)
        {
            // Example: limit processing to the first 1,000,000 rows
            return rowIndex < 1_000_000;
        }

        // Called after row properties are read; return true to process its cells
        public bool ProcessRow(Row row)
        {
            // No specific row-level logic needed here
            return true;
        }

        // Called for each cell column index; return true to read the cell
        public bool StartCell(int columnIndex)
        {
            return true;
        }

        // Called for each cell; here we simply output its address and value
        public bool ProcessCell(Cell cell)
        {
            Console.WriteLine($"Cell[{cell.Row},{cell.Column}] = {cell.Value}");
            return true;
        }
    }
}
