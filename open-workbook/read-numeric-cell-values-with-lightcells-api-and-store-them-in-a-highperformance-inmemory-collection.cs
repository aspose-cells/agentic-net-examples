// Title: C# – Extract numeric cells with Aspose.Cells LightCells API into a thread‑safe ConcurrentBag
// Description: Loads a large Excel workbook using Aspose.Cells, scans every worksheet for numeric cells, and pushes each double value into a ConcurrentBag<double> for fast, thread‑safe in‑memory storage. The sample also shows how to count the extracted cells and compute their sum.
// Keywords: Aspose.Cells LightCells numeric extraction | C# read Excel numeric values | ConcurrentBag<double> Excel data | high‑performance Excel parsing .NET | thread‑safe collection for Excel numbers | large workbook processing Aspose.Cells | numeric cell iteration C# | Excel to memory collection
// Common Searches: how to read only numeric cells from Excel with Aspose.Cells | store Excel numbers in a thread‑safe collection C# | Aspose.Cells LightCells API example for numeric data | fast in‑memory caching of Excel numeric values | C# extract and sum numeric cells from large workbook
// Developer Intent: Read every numeric cell from a workbook and keep the values in a high‑throughput, thread‑safe in‑memory collection for further processing.
// Use Cases: Compute aggregates (sum, average, min, max) on millions of numbers without re‑reading the file. | Feed numeric data into a parallel calculation engine for financial modeling or scientific simulations. | Cache spreadsheet numbers for instant lookup in downstream analytics or reporting modules.
// AI Prompts: Write Aspose.Cells LightCells code that streams only numeric cells into a ConcurrentBag<double> for maximum speed. | Show how to combine Parallel.ForEach with LightCells to extract numeric values from a massive Excel file into a thread‑safe collection. | Suggest memory‑optimisation techniques when storing tens of millions of double values extracted from Excel using Aspose.Cells.

using System;
using System.Collections.Concurrent;
using System.IO;
using Aspose.Cells;

namespace LightCellsNumericReader
{
    // Loads a large Excel workbook using Aspose.Cells, scans every worksheet for numeric cells, and pushes each double value into a ConcurrentBag<double> for fast, thread‑safe in‑memory storage. The sample also shows how to count the extracted cells and compute their sum.
    class Program
    {
        static void Main()
        {
            // Path to the large Excel file to be read
            string inputFile = "LargeData.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            try
            {
                // Load the workbook (standard mode)
                Workbook workbook = new Workbook(inputFile);

                // Thread‑safe collection for high‑performance in‑memory storage
                ConcurrentBag<double> numericValues = new ConcurrentBag<double>();

                // Iterate through each worksheet and its used cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell.Type == CellValueType.IsNumeric)
                            {
                                numericValues.Add(cell.DoubleValue);
                            }
                        }
                    }
                }

                // Output results
                Console.WriteLine($"Total numeric cells processed: {numericValues.Count}");

                double sum = 0;
                foreach (double val in numericValues)
                {
                    sum += val;
                }
                Console.WriteLine($"Sum of numeric values: {sum}");

                // Optionally, save the workbook (unchanged) to a new file
                workbook.Save("ProcessedOutput.xlsx");
            }
            catch (Exception ex)
            {
                // Runtime safety: capture and display any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
