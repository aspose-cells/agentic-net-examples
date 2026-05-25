using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsProgressDemo
{
    // Delegate for progress notification
    public delegate void ProgressChangedHandler(int processedCells);

    // Handles cell enumeration and raises progress events every 100 cells
    public class CellProcessor
    {
        // Event raised after each batch of 100 cells is processed
        public event ProgressChangedHandler ProgressChanged;

        // Enumerates all cells in the worksheet and raises progress events
        public void ProcessCells(Worksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            // Get the cells enumerator
            IEnumerator enumerator = worksheet.Cells.GetEnumerator();

            int processedCount = 0;

            while (enumerator.MoveNext())
            {
                // Access the current cell (optional processing can be added here)
                Cell cell = (Cell)enumerator.Current;
                // Example: just read the value to simulate work
                var value = cell.Value;

                processedCount++;

                // Every 100 cells, raise the progress event
                if (processedCount % 100 == 0)
                {
                    ProgressChanged?.Invoke(processedCount);
                }
            }

            // If the total number of cells is not a multiple of 100,
            // raise a final event to indicate completion
            if (processedCount % 100 != 0)
            {
                ProgressChanged?.Invoke(processedCount);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard Aspose.Cells creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (e.g., 1050 cells)
            int rows = 35;
            int cols = 30; // 35 * 30 = 1050 cells
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r}C{c}");
                }
            }

            // Instantiate the processor and subscribe to the progress event
            CellProcessor processor = new CellProcessor();
            processor.ProgressChanged += OnProgressChanged;

            // Process cells with progress reporting
            processor.ProcessCells(sheet);

            // Save the workbook (using the standard Aspose.Cells saving rule)
            workbook.Save("ProgressDemo.xlsx");
        }

        // Event handler that receives progress updates
        private static void OnProgressChanged(int processedCells)
        {
            Console.WriteLine($"Processed {processedCells} cells.");
        }
    }
}