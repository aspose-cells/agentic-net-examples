// Title: Report cell enumeration progress in Aspose.Cells by raising a batch event every 100 cells (C#)
// Description: The example defines a custom CellBatchProcessedEventArgs class and a CellProcessor that enumerates all cells of a worksheet using IEnumerator. After each hundred cells are read, the processor fires a BatchProcessed event with the cumulative count and also fires the event for any remaining cells at the end. A console program fills 1,050 cells, subscribes to the event, writes progress messages to the console, and optionally saves the workbook.
// Keywords: Aspose.Cells | C# | cell enumeration | progress event | batch processing | event handling | worksheet iteration | cell count reporting | large spreadsheet | console logging
// Common Searches: Aspose.Cells raise event after processing cells | C# progress callback while iterating worksheet cells | track cell processing progress in Aspose.Cells | batch processed event Aspose.Cells | enumerate cells with progress reporting Aspose.Cells
// Developer Intent: Add an event‑driven progress indicator that notifies after every 100 cells are processed during worksheet enumeration.
// Use Cases: Show real‑time progress in a UI (e.g., ProgressBar) while handling massive worksheets. | Log processed cell counts for auditing or debugging large spreadsheet operations. | Combine batch progress events with cancellation or throttling logic for long‑running tasks.
// AI Prompts: Generate C# code that iterates through an Aspose.Cells worksheet and raises a custom event after every 100 cells processed. | Create a Windows Forms example that subscribes to CellProcessor.BatchProcessed and updates a ProgressBar with the processed cell count. | Explain how to extend CellProcessor to accept a CancellationToken while still reporting batch progress events.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsProgressDemo
{
    // Event arguments containing the number of cells processed so far
    // The example defines a custom CellBatchProcessedEventArgs class and a CellProcessor that enumerates all cells of a worksheet using IEnumerator. After each hundred cells are read, the processor fires a BatchProcessed event with the cumulative count and also fires the event for any remaining cells at the end. A console program fills 1,050 cells, subscribes to the event, writes progress messages to the console, and optionally saves the workbook.
    public class CellBatchProcessedEventArgs : EventArgs
    {
        public int ProcessedCellCount { get; }

        public CellBatchProcessedEventArgs(int count)
        {
            ProcessedCellCount = count;
        }
    }

    // Processor that enumerates cells and raises an event after each 100 cells
    public class CellProcessor
    {
        // Event raised when a batch of cells has been processed
        public event EventHandler<CellBatchProcessedEventArgs> BatchProcessed;

        protected virtual void OnBatchProcessed(int count)
        {
            BatchProcessed?.Invoke(this, new CellBatchProcessedEventArgs(count));
        }

        // Enumerates all cells in the given worksheet
        public void ProcessCells(Worksheet worksheet)
        {
            IEnumerator enumerator = worksheet.Cells.GetEnumerator();
            int processed = 0;

            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Example processing: simply read the cell value
                var value = cell.Value;

                processed++;

                // Raise event after each hundred cells
                if (processed % 100 == 0)
                {
                    OnBatchProcessed(processed);
                }
            }

            // Raise event for the final partial batch (if any)
            if (processed % 100 != 0)
            {
                OnBatchProcessed(processed);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill 1050 cells (105 rows × 10 columns) with sample values
            int rows = 105;
            int cols = 10;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    sheet.Cells[i, j].PutValue($"R{i}C{j}");
                }
            }

            // Set up the processor and subscribe to the progress event
            CellProcessor processor = new CellProcessor();
            processor.BatchProcessed += Processor_BatchProcessed;

            // Process cells with progress reporting
            processor.ProcessCells(sheet);

            // Save the workbook (optional)
            workbook.Save("ProgressDemo.xlsx");
        }

        // Event handler that reports progress to the console
        private static void Processor_BatchProcessed(object sender, CellBatchProcessedEventArgs e)
        {
            Console.WriteLine($"Processed {e.ProcessedCellCount} cells.");
        }
    }
}
