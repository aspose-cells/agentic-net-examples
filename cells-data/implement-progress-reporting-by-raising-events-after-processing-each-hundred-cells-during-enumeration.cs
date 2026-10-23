// Title: Raise a progress event after every 100 cells while iterating an Aspose.Cells workbook in C#
// Common Searches: c# Aspose.Cells iterate through cells with progress callback every 100 cells | how to report processing progress during Excel cell enumeration using Aspose.Cells | Aspose.Cells raise custom event after processing a batch of cells in C# | track cell processing count with event handler while looping over used range in Aspose.Cells | C# example for progress notification every 100 cells in Aspose.Cells workbook
// Tags: Aspose.Cells cell enumeration progress event | C# batch cell processing callback | Excel workbook processing progress Aspose.Cells | event-driven cell iteration Aspose.Cells | progress reporting for used range cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsProgressDemo
{
    // Delegate for progress reporting
    public delegate void ProgressEventHandler(int processedCells);

    // The sample defines a WorkbookProcessor that loads an Excel file with Aspose.Cells, iterates over all cells in the first worksheet, and raises a ProgressChanged event after each 100 cells (including any remaining cells). Consumers subscribe to the event to receive the processed cell count, enabling real‑time progress reporting during workbook processing.
    public class WorkbookProcessor
    {
        // Event raised after each 100 cells are processed
        public event ProgressEventHandler ProgressChanged;

        // Method to process the workbook
        public void Process(string inputFilePath)
        {
            // Load the workbook (using the provided load rule)
            Workbook workbook = new Workbook(inputFilePath);

            // Get the first worksheet (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            int processedCount = 0;

            // Enumerate all cells in the used range
            foreach (Cell cell in cells)
            {
                // Example processing: read the cell value (replace with actual logic)
                var value = cell.Value;

                // Increment processed counter
                processedCount++;

                // After each 100 cells, raise the progress event
                if (processedCount % 100 == 0)
                {
                    OnProgressChanged(processedCount);
                }
            }

            // Optionally raise event for any remaining cells (<100)
            if (processedCount % 100 != 0)
            {
                OnProgressChanged(processedCount);
            }

            // Save the workbook (using the provided save rule)
            workbook.Save("Processed_" + System.IO.Path.GetFileName(inputFilePath));
        }

        // Helper method to invoke the event safely
        protected virtual void OnProgressChanged(int processedCells)
        {
            ProgressChanged?.Invoke(processedCells);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage
            string inputPath = "Sample.xlsx";

            WorkbookProcessor processor = new WorkbookProcessor();

            // Subscribe to the progress event
            processor.ProgressChanged += Processor_ProgressChanged;

            // Start processing
            processor.Process(inputPath);
        }

        // Event handler that receives progress updates
        private static void Processor_ProgressChanged(int processedCells)
        {
            Console.WriteLine($"Processed {processedCells} cells.");
        }
    }
}
