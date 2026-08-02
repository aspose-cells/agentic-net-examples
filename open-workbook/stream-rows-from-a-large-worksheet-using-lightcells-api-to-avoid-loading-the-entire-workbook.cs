// Title: Stream rows from a large Excel worksheet with Aspose.Cells LightCells API (C#)
// Description: Demonstrates how to use a custom LightCellsDataHandler with LoadOptions to read and process an Excel workbook row‑by‑row, logging each sheet, row, and cell while keeping memory usage low, and optionally saving the workbook after processing.
// Keywords: Aspose.Cells | LightCells API | C# streaming rows | large Excel workbook | memory‑efficient Excel processing | LoadOptions LightCellsDataHandler | row‑by‑row Excel read | GitHub example | open workbook without full load | Excel data extraction C#
// Common Searches: Aspose.Cells LightCells stream rows C# | process large Excel file without loading whole workbook | C# LightCellsDataHandler example | read Excel rows one at a time Aspose | memory efficient Excel parsing Aspose.Cells
// Developer Intent: Read and manipulate rows of a massive worksheet without loading the entire workbook into memory, using Aspose.Cells LightCells streaming capabilities.
// Use Cases: Migrate data from a multi‑gigabyte Excel file to a database by streaming rows. | Audit or log every cell value in a large workbook without exhausting RAM. | Apply row‑level transformations (e.g., format changes, calculations) and save the result while keeping the process lightweight.
// AI Prompts: Create a LightCellsDataHandler that skips rows where column A is empty. | Show how to write each processed row to a new workbook while streaming with LightCells. | Add comprehensive error handling for missing cells and type conversion errors during LightCells row streaming.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamingExample
{
    // Demonstrates how to use a custom LightCellsDataHandler with LoadOptions to read and process an Excel workbook row‑by‑row, logging each sheet, row, and cell while keeping memory usage low, and optionally saving the workbook after processing.
    class Program
    {
        static void Main()
        {
            // Path to the large workbook to be processed
            string inputPath = "LargeFile_original.xlsx";

            // Path where the processed workbook will be saved (optional)
            string outputPath = "ProcessedLargeFile.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    // Create a minimal placeholder workbook
                    var placeholder = new Workbook();
                    placeholder.Worksheets[0].Name = "Sheet1";
                    placeholder.Save(inputPath);
                    Console.WriteLine($"Created placeholder workbook at {inputPath}");
                }

                // Create an instance of the custom LightCellsDataHandler
                var handler = new StreamingHandler();

                // Configure load options to use the LightCellsDataHandler
                var loadOptions = new LoadOptions
                {
                    LightCellsDataHandler = handler
                };

                // Load the workbook in light cells mode – rows are streamed and processed
                var workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook after processing (optional, demonstrates full lifecycle)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Custom handler that streams rows and cells without loading the whole workbook into memory
    class StreamingHandler : LightCellsDataHandler
    {
        // Called before reading a worksheet; return true to process this sheet
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Starting to process sheet: {sheet.Name}");
            return true;
        }

        // Called before reading each row; return true to process the row
        public bool StartRow(int rowIndex)
        {
            return true; // Process every row
        }

        // Called after the row header is read; return true to process its cells
        public bool ProcessRow(Row row)
        {
            Console.WriteLine($"Processing Row {row.Index}");
            return true; // Continue to cell processing for this row
        }

        // Called before reading each cell in the current row; return true to process the cell
        public bool StartCell(int columnIndex)
        {
            return true; // Process every cell in the row
        }

        // Called for each cell that is to be processed
        public bool ProcessCell(Cell cell)
        {
            Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
            return true; // Continue processing
        }
    }
}
