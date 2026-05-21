using System;
using Aspose.Cells;

namespace LightCellsCustomBufferDemo
{
    // Custom handler to process cells while reading in LightCells mode
    public class CustomLightCellsDataHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet)
        {
            // Process all sheets
            return true;
        }

        public bool StartRow(int rowIndex)
        {
            // Process all rows
            return true;
        }

        public bool ProcessRow(Row row)
        {
            // No special row processing needed
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            // Process all cells
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            // Example: simply output cell address and value
            Console.WriteLine($"Processing {cell.Name}: {cell.Value}");
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to a large Excel file (replace with actual file path)
            string inputPath = "LargeWorkbook.xlsx";
            string outputPath = "ProcessedWorkbook.xlsx";

            // Create a custom LightCellsDataHandler
            var dataHandler = new CustomLightCellsDataHandler();

            // Configure load options to use LightCells mode with the custom handler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = dataHandler;

            // Load the workbook using LightCells (streaming) mode
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet's cells
            Cells cells = workbook.Worksheets[0].Cells;

            // Configure memory usage to FileCache to reduce memory pressure.
            // This effectively uses a temporary file as a buffer, which can be tuned by the OS.
            cells.MemorySetting = MemorySetting.FileCache;

            // Enable multi‑thread reading for high‑throughput scenarios.
            // Note: This may degrade single‑thread performance but improves parallel access.
            cells.MultiThreadReading = true;

            // Save the processed workbook (standard save; LightCells mode is only for reading)
            workbook.Save(outputPath);
        }
    }
}