// Title: Compare initialized cell count before and after removing formatting‑only cells using Aspose.Cells for .NET
// Common Searches: how to count initialized cells in an Excel workbook with Aspose.Cells C# | remove formatting‑only cells in a .xlsx file using Aspose.Cells and get cell count reduction | Aspose.Cells C# example to reset cell style to default and compare cell counts | measure reduction of initialized cells after clearing formatting with Aspose.Cells | C# code to iterate over worksheet cells and count those with values or custom styles using Aspose.Cells
// Tags: initialized cell counting Aspose.Cells | clear formatting‑only cells Aspose.Cells | reset cell style to default Aspose.Cells | worksheet cell count reduction Aspose.Cells | C# iterate over cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, counts cells that are initialized (contain a value or a non‑default style), resets the style of cells that have only formatting, recounts initialized cells, displays the before/after counts and the reduction, and optionally saves the cleaned workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Create a default style once to reuse for comparisons
                Style defaultStyle = workbook.CreateStyle();

                // ---------- Count initialized cells BEFORE removing formatting‑only cells ----------
                int initializedBefore = 0;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Consider a cell initialized if it has a value or a non‑default style
                        if (cell.Value != null || !cell.GetStyle().Equals(defaultStyle))
                        {
                            initializedBefore++;
                        }
                    }
                }

                // ---------- Remove cells that contain only formatting ----------
                try
                {
                    foreach (Cell cell in cells)
                    {
                        // If the cell has no value but has custom formatting, reset its style
                        if (cell.Value == null && !cell.GetStyle().Equals(defaultStyle))
                        {
                            cell.SetStyle(defaultStyle);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to remove formatting‑only cells. {ex.Message}");
                }

                // ---------- Count initialized cells AFTER removal ----------
                int initializedAfter = 0;
                maxRow = cells.MaxDataRow;
                maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Value != null || !cell.GetStyle().Equals(defaultStyle))
                        {
                            initializedAfter++;
                        }
                    }
                }

                // Output the comparison result
                Console.WriteLine($"Initialized cells before removal: {initializedBefore}");
                Console.WriteLine($"Initialized cells after removal : {initializedAfter}");
                Console.WriteLine($"Reduction achieved: {initializedBefore - initializedAfter}");

                // Save the modified workbook (optional)
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Modified workbook saved as \"{outputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
