// Title: Raise a custom progress event after every 100 cells while enumerating a worksheet with Aspose.Cells for .NET
// AI Prompts: Write a C# CellProcessor class that iterates through all cells of an Aspose.Cells worksheet and triggers a CellsBatchProcessed event after each 100‑cell batch. | Add a configurable batchSize argument to the cell enumeration method so that a progress event is raised after the specified number of cells have been processed. | Create an event handler that logs the processed cell count each time the CellsBatchProcessed event fires during worksheet traversal.
// Common Searches: Aspose.Cells C# raise progress event every 100 cells during worksheet iteration | how to implement batch progress notifications while looping through cells in Aspose.Cells | C# example of custom event for cell processing count in Aspose.Cells workbook | track enumeration progress of Excel cells using Aspose.Cells .NET API
// Tags: Aspose.Cells cell enumeration progress event | C# batch processing of worksheet cells | custom progress event Aspose.Cells | enumerate Excel cells with callback | cell batch size configurable Aspose

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsProgressDemo
{
    // Delegate for progress event
    public delegate void CellsBatchProcessedHandler(object sender, CellsBatchProcessedEventArgs e);

    // Event arguments containing the number of cells processed so far
    // The sample defines a CellsBatchProcessed event that fires after each 100‑cell batch during worksheet enumeration. A CellProcessor class handles the iteration, raises the event, and a console handler displays the cumulative cell count, demonstrating progress reporting in Aspose.Cells for .NET.
    public class CellsBatchProcessedEventArgs : EventArgs
    {
        public int CellsProcessed { get; }

        public CellsBatchProcessedEventArgs(int cellsProcessed)
        {
            CellsProcessed = cellsProcessed;
        }
    }

    // Processor that enumerates cells and raises an event after each 100 cells
    public class CellProcessor
    {
        // Event raised after each batch of 100 cells
        public event CellsBatchProcessedHandler CellsBatchProcessed;

        // Enumerates all cells in the given worksheet
        public void ProcessCells(Worksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            IEnumerator enumerator = worksheet.Cells.GetEnumerator();
            int count = 0;

            while (enumerator.MoveNext())
            {
                // Access the current cell (optional processing can be added here)
                Cell cell = (Cell)enumerator.Current;

                // Example processing: just read the value (no modification)
                var value = cell.Value;

                count++;

                // After each 100 cells, raise the progress event
                if (count % 100 == 0)
                {
                    OnCellsBatchProcessed(new CellsBatchProcessedEventArgs(count));
                }
            }

            // Raise event for any remaining cells that didn't fill a complete batch
            if (count % 100 != 0)
            {
                OnCellsBatchProcessed(new CellsBatchProcessedEventArgs(count));
            }
        }

        // Helper method to invoke the event safely
        protected virtual void OnCellsBatchProcessed(CellsBatchProcessedEventArgs e)
        {
            CellsBatchProcessed?.Invoke(this, e);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (more than 200 cells to demonstrate batching)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 15; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Instantiate the processor and subscribe to the progress event
            CellProcessor processor = new CellProcessor();
            processor.CellsBatchProcessed += Processor_CellsBatchProcessed;

            // Process cells with progress reporting
            processor.ProcessCells(worksheet);

            // Save the workbook (using the standard save method)
            workbook.Save("ProgressDemo.xlsx");
        }

        // Event handler that receives progress notifications
        private static void Processor_CellsBatchProcessed(object sender, CellsBatchProcessedEventArgs e)
        {
            Console.WriteLine($"Processed {e.CellsProcessed} cells.");
        }
    }
}
