// Title: Report progress every 100 cells while iterating a worksheet with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells raise custom event after each hundred cells processed | track cell processing progress during worksheet enumeration with Aspose.Cells | how to implement progress callback while looping through Excel cells in .NET | subscribe to progress events when modifying cells using Aspose.Cells | report number of processed cells every 100 iterations in Aspose.Cells C#
// Tags: Aspose.Cells progress event handling | enumerate worksheet cells Aspose.Cells | cell processing loop with progress callback | save workbook after cell iteration Aspose.Cells | C# custom progress reporting for Excel

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsProgressDemo
{
    // Delegate for progress reporting
    public delegate void ProgressEventHandler(object sender, ProgressEventArgs e);

    // Event arguments containing the number of cells processed
    // Demonstrates how to raise a custom ProgressChanged event after every 100 cells are processed while enumerating all used cells in the first worksheet of an Excel file using Aspose.Cells for .NET, includes event subscription, placeholder cell manipulation, and saving the updated workbook.
    public class ProgressEventArgs : EventArgs
    {
        public int CellsProcessed { get; }

        public ProgressEventArgs(int cellsProcessed)
        {
            CellsProcessed = cellsProcessed;
        }
    }

    public class CellProcessor
    {
        // Event raised after each 100 cells are processed
        public event ProgressEventHandler? ProgressChanged;

        // Method to raise the ProgressChanged event
        protected virtual void OnProgressChanged(int cellsProcessed)
        {
            ProgressChanged?.Invoke(this, new ProgressEventArgs(cellsProcessed));
        }

        // Core logic: enumerate all cells in the first worksheet and raise progress events
        public void Process(string inputFilePath, string outputFilePath)
        {
            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            Workbook? workbook = null;
            try
            {
                // Load the workbook (no password supplied)
                workbook = new Workbook(inputFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                int processedCount = 0;

                // Enumerate through all used cells
                foreach (Cell cell in cells)
                {
                    // Example processing: convert the cell value to string and back (placeholder for real work)
                    string value = cell.StringValue;
                    cell.PutValue(value);

                    processedCount++;

                    // After each hundred cells, raise the progress event
                    if (processedCount % 100 == 0)
                    {
                        OnProgressChanged(processedCount);
                    }
                }

                // Save the workbook
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during processing or saving: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Paths can be replaced with actual file locations
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            CellProcessor processor = new CellProcessor();

            // Subscribe to the progress event
            processor.ProgressChanged += Processor_ProgressChanged;

            // Start processing
            processor.Process(inputPath, outputPath);
        }

        // Event handler that receives progress updates
        private static void Processor_ProgressChanged(object sender, ProgressEventArgs e)
        {
            Console.WriteLine($"Processed {e.CellsProcessed} cells.");
        }
    }
}
