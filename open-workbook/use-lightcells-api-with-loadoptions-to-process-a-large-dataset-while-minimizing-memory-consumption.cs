using System;
using Aspose.Cells;

namespace LightCellsProcessingDemo
{
    // Author: Aspose.Cells .NET example author
    // Custom handler that processes cells in a memory‑efficient way using LightCells API
    public class CustomLightCellsDataHandler : LightCellsDataHandler
    {
        // Called when a worksheet starts processing
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Processing sheet: {sheet.Name}");
            return true; // continue processing this sheet
        }

        // Called before a row is processed
        public bool StartRow(int rowIndex)
        {
            // Optionally, you could skip rows by returning false
            return true; // continue processing this row
        }

        // Called after a row has been processed (optional implementation)
        public bool ProcessRow(Row row)
        {
            // No custom row‑level logic needed for this demo
            return true;
        }

        // Called before a cell in the current row is processed
        public bool StartCell(int columnIndex)
        {
            return true; // continue processing this cell
        }

        // Called for each cell; here we simply output its address and value
        public bool ProcessCell(Cell cell)
        {
            Console.WriteLine($"Cell[{cell.Row},{cell.Column}] = {cell.Value}");
            return true; // continue processing subsequent cells
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the large workbook to be processed
            const string inputPath = "LargeDataset.xlsx";
            const string outputPath = "ProcessedLargeDataset.xlsx";

            // Create an instance of the custom LightCellsDataHandler
            var dataHandler = new CustomLightCellsDataHandler();

            // Configure LoadOptions to use the LightCellsDataHandler
            var loadOptions = new LoadOptions
            {
                LightCellsDataHandler = dataHandler,
                // Optional: reduce memory footprint by disabling unnecessary features
                KeepUnparsedData = false,
                AutoFitterOptions = null
            };

            // Load the workbook in light‑weight mode; only cell data is streamed through the handler
            using (var workbook = new Workbook(inputPath, loadOptions))
            {
                // At this point the handler has processed the workbook contents.
                // If you need to create a new workbook based on processed data,
                // you could populate it here using a LightCellsDataProvider (not shown).

                // Save the (potentially unchanged) workbook to demonstrate the workflow
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
        }
    }
}