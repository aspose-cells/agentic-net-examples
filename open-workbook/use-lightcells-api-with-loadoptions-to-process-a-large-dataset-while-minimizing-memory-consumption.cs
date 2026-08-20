// Title: Low‑Memory Processing of Large Excel Files with Aspose.Cells LightCells and LoadOptions (C#)
// Description: Demonstrates how to stream a massive Excel workbook using Aspose.Cells LightCellsDataHandler with LoadOptions. The custom handler prints each cell, accumulates a numeric sum, and keeps memory usage minimal by disabling KeepUnparsedData. The workbook is then saved without loading the full file into memory.
// Keywords: Aspose.Cells LightCells | C# LightCellsDataHandler | LoadOptions KeepUnparsedData false | stream large Excel file | low memory Excel processing | calculate sum while streaming | memory‑efficient workbook loading | Aspose.Cells large dataset
// Common Searches: Aspose.Cells LightCells example for large worksheets | How to reduce memory usage with LoadOptions in Aspose.Cells | Stream Excel cells in C# without loading entire workbook | Calculate numeric sum using LightCellsDataHandler | Disable KeepUnparsedData to save memory Aspose.Cells
// Developer Intent: The developer needs to process a huge Excel workbook in a streaming fashion, compute aggregates on‑the‑fly, and keep RAM consumption as low as possible.
// Use Cases: Read and log every cell of a multi‑gigabyte workbook without full in‑memory load. | Aggregate numeric columns (e.g., totals, averages) while streaming data. | Perform read‑only analysis on large spreadsheets and optionally save the unchanged file.
// AI Prompts: Create a LightCellsDataHandler that writes each processed cell to a CSV file while maintaining low memory usage. | Modify the handler to skip rows where a specific column value meets a condition, still using LightCells. | Provide LoadOptions settings for optimal performance when processing a 10 GB Excel file with LightCells.

using System;
using Aspose.Cells;

namespace LightCellsProcessingDemo
{
    // Custom handler that processes cells in a streaming (lightweight) manner.
    // This implementation simply prints each cell value and accumulates a numeric sum.
    // Demonstrates how to stream a massive Excel workbook using Aspose.Cells LightCellsDataHandler with LoadOptions. The custom handler prints each cell, accumulates a numeric sum, and keeps memory usage minimal by disabling KeepUnparsedData. The workbook is then saved without loading the full file into memory.
    public class SummingLightCellsHandler : LightCellsDataHandler
    {
        private double _numericSum = 0;

        // Called when a worksheet is about to be processed.
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Processing sheet: {sheet.Name}");
            // Return true to continue processing this sheet.
            return true;
        }

        // Called before a row is processed.
        public bool StartRow(int rowIndex)
        {
            // Return true to process the row.
            return true;
        }

        // Called after the row object is created; can be used to inspect row properties.
        public bool ProcessRow(Row row)
        {
            // Return true to allow processing of the cells in this row.
            return true;
        }

        // Called before a cell in the current row is processed.
        public bool StartCell(int columnIndex)
        {
            // Return true to process the cell.
            return true;
        }

        // Called for each cell that needs to be processed.
        public bool ProcessCell(Cell cell)
        {
            // Output cell address and value.
            Console.WriteLine($"Cell[{cell.Row},{cell.Column}] = {cell.Value}");

            // If the cell contains a numeric value, add it to the running sum.
            if (cell.Type == CellValueType.IsNumeric)
            {
                _numericSum += cell.DoubleValue;
            }

            // Continue processing subsequent cells.
            return true;
        }

        // Expose the accumulated sum after processing.
        public double GetNumericSum()
        {
            return _numericSum;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the large Excel file to be processed.
            const string inputPath = "LargeDataFile.xlsx";
            const string outputPath = "ProcessedLargeDataFile.xlsx";

            // Create load options and assign the custom LightCellsDataHandler.
            LoadOptions loadOptions = new LoadOptions();
            var handler = new SummingLightCellsHandler();
            loadOptions.LightCellsDataHandler = handler;

            // Disable keeping unparsed data to further reduce memory usage.
            loadOptions.KeepUnparsedData = false;

            // Load the workbook using the LightCells mode.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // After loading, retrieve the numeric sum calculated during streaming.
            Console.WriteLine($"Total numeric sum of processed cells: {handler.GetNumericSum()}");

            // Save the workbook (even if unchanged) using the standard save method.
            workbook.Save(outputPath);
        }
    }
}
