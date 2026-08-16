// Title: C# – Apply a 10‑Row Moving Average to Every Column with Aspose.Cells SetSharedFormula
// Description: Creates a workbook, fills a 30×3 range, defines a 10‑row sliding window, and uses SetSharedFormula with AVERAGE to compute rolling averages for each column, then calculates and saves the file.
// Keywords: Aspose.Cells | C# | .NET | moving average | sliding window | rolling average | SetSharedFormula | AVERAGE formula | Excel automation | time‑series analysis
// Common Searches: Aspose.Cells 10‑row moving average C# | SetSharedFormula rolling average example | how to calculate sliding window average with Aspose.Cells | C# Excel moving average using shared formula | Aspose.Cells time series average calculation
// Developer Intent: Generate a worksheet that automatically computes a 10‑row moving average for each column using a shared formula.
// Use Cases: Sales dashboards that display a 10‑day rolling average per product. | Sensor data analysis where each column needs a sliding‑window average. | Financial models that require moving averages across multiple series.
// AI Prompts: Write C# code with Aspose.Cells to apply a 10‑row moving average to all columns using SetSharedFormula. | Modify the example to accept a custom window size and automatically detect the number of columns. | Show how to read the calculated moving‑average values after workbook.CalculateFormula().

using System;
using Aspose.Cells;

namespace MovingAverageExample
{
    // Creates a workbook, fills a 30×3 range, defines a 10‑row sliding window, and uses SetSharedFormula with AVERAGE to compute rolling averages for each column, then calculates and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the size of the data range
            int totalRows = 30;      // total number of rows with data
            int totalCols = 3;       // total number of columns to process
            int windowSize = 10;     // sliding window for moving average

            // Populate the worksheet with sample numeric data
            // (you can replace this with your own data loading logic)
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Example: value = row index + 1 (so rows start at 1)
                    worksheet.Cells[row, col].PutValue(row + 1);
                }
            }

            // The first moving‑average result appears in the row that completes the first window
            int firstAvgRowIndex = windowSize - 1;                     // zero‑based index
            int avgResultRows = totalRows - windowSize + 1;            // how many average values we will have

            // Apply the moving‑average formula to each column using a shared formula
            for (int col = 0; col < totalCols; col++)
            {
                // Convert column index to Excel column letter (A, B, C, …)
                string colLetter = CellsHelper.ColumnIndexToName(col);

                // Formula for the first cell of the average range.
                // Relative references will shift automatically when the formula is copied down.
                string formula = $"=AVERAGE({colLetter}1:{colLetter}{windowSize})";

                // Starting cell where the first average will be placed
                Cell startCell = worksheet.Cells[firstAvgRowIndex, col];

                // Set a shared formula that will fill the column downwards.
                // Parameters: shared formula, number of rows to fill, number of columns (1).
                startCell.SetSharedFormula(formula, avgResultRows, 1);
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook with the moving‑average results
            workbook.Save("MovingAverage.xlsx");
        }
    }
}
