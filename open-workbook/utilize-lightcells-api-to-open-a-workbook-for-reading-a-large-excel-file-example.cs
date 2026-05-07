using System;
using Aspose.Cells;

namespace LightCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the large Excel file to be processed
            string inputPath = "LargeFile_original.xlsx";

            // Path where the processed workbook will be saved (optional)
            string outputPath = "ProcessedLargeFile.xlsx";

            // Create an instance of the custom LightCellsDataHandler
            var handler = new CustomLightCellsDataHandler();

            // Configure LoadOptions to use the LightCellsDataHandler
            var loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = handler;

            // Load the workbook in LightCells mode for efficient streaming processing
            var workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook after processing (optional, demonstrates the lifecycle)
            workbook.Save(outputPath);
        }
    }

    // Custom implementation of LightCellsDataHandler for streaming read
    public class CustomLightCellsDataHandler : LightCellsDataHandler
    {
        // Called when a worksheet starts processing
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Starting to process sheet: {sheet.Name}");
            return true; // Continue processing this sheet
        }

        // Called before a row is processed
        public bool StartRow(int rowIndex)
        {
            // Return true to process the row
            return true;
        }

        // Called after row properties are read
        public bool ProcessRow(Row row)
        {
            Console.WriteLine($"Processing row: {row.Index}");
            return true; // Continue to cell processing within this row
        }

        // Called before a cell in the current row is processed
        public bool StartCell(int columnIndex)
        {
            // Return true to process the cell
            return true;
        }

        // Called to process the actual cell data
        public bool ProcessCell(Cell cell)
        {
            // Example: output cell address and its string representation
            Console.WriteLine($"Cell {cell.Name}: {cell.StringValue}");
            return true; // Continue processing subsequent cells
        }
    }
}