// Title: C# – Monitor Memory While Loading Large Excel Files with Aspose.Cells LightCells API
// Description: Demonstrates how to extend LightCellsDataHandler with a custom MemoryMonitoringHandler that logs managed (GC.GetTotalMemory) and private (Process.PrivateMemorySize64) memory at each sheet, row, and cell step. The example sets MemorySetting.MemoryPreference, loads a large workbook in LightCells mode, and saves the result, providing a clear pattern for profiling memory consumption during high‑volume Excel processing.
// Keywords: Aspose.Cells | LightCells API | C# memory monitoring | large Excel workbook | GC.GetTotalMemory | Process.PrivateMemorySize64 | MemorySetting.MemoryPreference | Excel performance profiling | LightCellsDataHandler example | low‑memory Excel processing
// Common Searches: how to track memory usage with Aspose.Cells LightCells | lightcells memory monitoring C# example | log managed and private memory during Excel processing | reduce memory consumption when loading big XLSX files | Aspose.Cells MemorySetting.MemoryPreference usage
// Developer Intent: The developer needs to process a massive Excel workbook using LightCells and capture detailed memory statistics at each processing stage to identify and mitigate memory spikes.
// Use Cases: Profile memory consumption of each sheet, row, and cell while loading a large workbook in LightCells mode. | Compare memory footprints between default loading and MemorySetting.MemoryPreference for large files. | Validate that the LightCells workflow stays within memory limits before saving the processed workbook.
// AI Prompts: Generate a C# version of the MemoryMonitoringHandler that writes memory metrics to a CSV file instead of the console. | Show how to integrate Windows Performance Counters with the handler to record peak memory and CPU usage during workbook processing. | Explain how to modify the handler to abort processing of rows when memory usage exceeds a configurable threshold.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace LightCellsMemoryMonitoring
{
    // Custom handler that processes cells using LightCells API
    // and reports memory usage statistics during processing.
    // Demonstrates how to extend LightCellsDataHandler with a custom MemoryMonitoringHandler that logs managed (GC.GetTotalMemory) and private (Process.PrivateMemorySize64) memory at each sheet, row, and cell step. The example sets MemorySetting.MemoryPreference, loads a large workbook in LightCells mode, and saves the result, providing a clear pattern for profiling memory consumption during high‑volume Excel processing.
    public class MemoryMonitoringHandler : LightCellsDataHandler
    {
        // Called when a worksheet starts to be processed.
        public bool StartSheet(Worksheet sheet)
        {
            ReportMemory("StartSheet", sheet.Name);
            // Continue processing this sheet.
            return true;
        }

        // Called before a row is processed.
        public bool StartRow(int rowIndex)
        {
            ReportMemory("StartRow", $"Row {rowIndex}");
            // Continue processing this row.
            return true;
        }

        // Called after a row object is created; can be used to read row properties.
        public bool ProcessRow(Row row)
        {
            ReportMemory("ProcessRow", $"Row {row.Index}");
            // Continue processing cells in this row.
            return true;
        }

        // Called before a cell is processed.
        public bool StartCell(int columnIndex)
        {
            ReportMemory("StartCell", $"Column {columnIndex}");
            // Continue processing this cell.
            return true;
        }

        // Called after a cell object is created; can be used to read cell data.
        public bool ProcessCell(Cell cell)
        {
            ReportMemory("ProcessCell", $"Cell {cell.Name}");
            // Example: just output the cell value (optional).
            // Console.WriteLine($"Value: {cell.Value}");
            return true;
        }

        // Helper method to output memory usage.
        private void ReportMemory(string stage, string context)
        {
            // Get managed memory used by the CLR.
            long managedBytes = GC.GetTotalMemory(forceFullCollection: false);
            // Get total private memory used by the process.
            long privateBytes = Process.GetCurrentProcess().PrivateMemorySize64;

            Console.WriteLine($"{stage} - {context}: Managed = {managedBytes / (1024 * 1024)} MB, " +
                              $"Private = {privateBytes / (1024 * 1024)} MB");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the large workbook to be processed.
            string inputPath = "LargeFile_original.xlsx";
            // Path where the processed workbook will be saved.
            string outputPath = "ProcessedLargeFile.xlsx";

            // Create load options and assign the custom LightCellsDataHandler.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new MemoryMonitoringHandler();

            // Optional: set memory mode to reduce memory consumption while loading.
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;

            // Load the workbook using LightCells mode.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // At this point the handler has already processed the workbook
            // and reported memory usage for each step.

            // Save the workbook to demonstrate the complete workflow.
            workbook.Save(outputPath);

            Console.WriteLine("Processing completed. Workbook saved to: " + outputPath);
        }
    }
}
