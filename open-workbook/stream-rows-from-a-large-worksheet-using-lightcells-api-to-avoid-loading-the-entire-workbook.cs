using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsStreamingExample
{
    // Custom handler that streams rows and cells without loading the whole workbook into memory
    public class RowStreamingHandler : LightCellsDataHandler
    {
        // Called before processing a worksheet. Return true to process this sheet.
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Processing sheet: {sheet.Name}");
            return true;
        }

        // Called before each row. Return true to process the row.
        public bool StartRow(int rowIndex)
        {
            // You can filter rows here if needed.
            return true;
        }

        // Called after a row is prepared. Return true if its cells should be processed.
        public bool ProcessRow(Row row)
        {
            // Example: output row index.
            Console.WriteLine($"Row {row.Index} started.");
            return true; // Continue to process cells in this row.
        }

        // Called before each cell in the current row. Return true to process the cell.
        public bool StartCell(int columnIndex)
        {
            return true;
        }

        // Called for each cell that was accepted by StartCell.
        public bool ProcessCell(Cell cell)
        {
            // Output cell address and its value.
            Console.WriteLine($"  Cell {cell.Name}: {cell.Value}");
            return true; // Continue processing.
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the large Excel file.
            string inputPath = "LargeFile.xlsx";

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Create load options and assign the custom LightCellsDataHandler.
                LoadOptions loadOptions = new LoadOptions
                {
                    LightCellsDataHandler = new RowStreamingHandler()
                };

                // Load the workbook in LightCells mode. The handler will stream rows/cells.
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Optionally, save the workbook after processing (uses default saving).
                // This demonstrates the save lifecycle without custom provider.
                string outputPath = "ProcessedLargeFile.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}