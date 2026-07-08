using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsProgressDemo
{
    // Event arguments containing the number of cells processed so far
    public class CellsProcessedEventArgs : EventArgs
    {
        public int ProcessedCount { get; }

        public CellsProcessedEventArgs(int processedCount)
        {
            ProcessedCount = processedCount;
        }
    }

    // Processor that enumerates cells and raises an event after each 100 cells
    public class CellProcessor
    {
        // Event raised after each batch of 100 cells
        public event EventHandler<CellsProcessedEventArgs> CellsBatchProcessed;

        // Helper to invoke the event safely
        protected virtual void OnCellsBatchProcessed(int count)
        {
            CellsBatchProcessed?.Invoke(this, new CellsProcessedEventArgs(count));
        }

        // Enumerates all cells in the given worksheet and reports progress
        public void ProcessCells(Worksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            IEnumerator enumerator = worksheet.Cells.GetEnumerator();
            int processed = 0;

            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Example processing: just read the value (could be any logic)
                var value = cell.Value;

                processed++;

                // Raise event after each 100 cells
                if (processed % 100 == 0)
                {
                    OnCellsBatchProcessed(processed);
                }
            }

            // If total cells is not a multiple of 100, raise a final event
            if (processed % 100 != 0)
            {
                OnCellsBatchProcessed(processed);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (e.g., 550 cells)
            int rows = 55;
            int cols = 10;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r}C{c}");
                }
            }

            // Instantiate the processor and subscribe to the progress event
            CellProcessor processor = new CellProcessor();
            processor.CellsBatchProcessed += Processor_CellsBatchProcessed;

            // Process cells with progress reporting
            processor.ProcessCells(sheet);

            // Save the workbook
            workbook.Save("ProgressDemo.xlsx");
        }

        // Event handler that receives progress notifications
        private static void Processor_CellsBatchProcessed(object sender, CellsProcessedEventArgs e)
        {
            Console.WriteLine($"Processed {e.ProcessedCount} cells.");
        }
    }
}