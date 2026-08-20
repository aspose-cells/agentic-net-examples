// Title: Aspose.Cells .NET: Compare instantiated cell count before & after removing formatting‑only cells
// Description: Shows how to create a workbook, add value cells and formatting‑only cells, capture the total instantiated cells with Cells.CountLarge, clear styles from empty cells, call Workbook.RemoveUnusedStyles, and display the reduction in cell count to evaluate memory savings.
// Keywords: Aspose.Cells | .NET | CountLarge | RemoveUnusedStyles | formatting only cells | cell count reduction | memory optimization | worksheet cleanup | C# example | instantiated cells
// Common Searches: Aspose.Cells count instantiated cells | remove formatting only cells Aspose.Cells .NET | how to reduce worksheet memory Aspose.Cells | CountLarge before after RemoveUnusedStyles | C# Aspose.Cells clear empty cell styles | measure cell count reduction Aspose.Cells
// Developer Intent: Determine how many Cell objects are eliminated by clearing formatting‑only cells and removing unused styles in an Aspose.Cells workbook.
// Use Cases: Assess memory impact of styled empty cells in large spreadsheets. | Validate performance gains after workbook cleanup in automated processing pipelines. | Generate reports on cell count reduction for optimization decisions. | Ensure generated Excel files stay within size limits for .NET applications.
// AI Prompts: Generate C# code using Aspose.Cells to identify empty cells with non‑default styles and reset them to the default style. | Create a method that returns the percentage reduction of Cells.CountLarge after invoking RemoveUnusedStyles. | Explain why formatting‑only cells increase Cells.CountLarge and how RemoveUnusedStyles improves memory usage.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add value cells and formatting‑only cells, capture the total instantiated cells with Cells.CountLarge, clear styles from empty cells, call Workbook.RemoveUnusedStyles, and display the reduction in cell count to evaluate memory savings.
    public class FormattingOnlyCellsReductionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate the worksheet:
                // - Cells with values
                // - Cells with only formatting (no value)
                cells["A1"].PutValue("Data 1");
                cells["A2"].PutValue("Data 2");

                // Cell B1: formatting only (bold font)
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;
                cells["B1"].SetStyle(boldStyle);

                // Cell B2: formatting only (red background)
                Style redBgStyle = workbook.CreateStyle();
                redBgStyle.ForegroundColor = System.Drawing.Color.Red;
                redBgStyle.Pattern = BackgroundType.Solid;
                cells["B2"].SetStyle(redBgStyle);

                // Cell C1: both value and formatting
                cells["C1"].PutValue("Data 3");
                cells["C1"].SetStyle(boldStyle);

                // Count of instantiated Cell objects before cleanup
                long countBefore = cells.CountLarge;
                Console.WriteLine($"Instantiated cells before removing formatting‑only cells: {countBefore}");

                // Remove formatting from cells that have no value
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // If the cell has no value but has a non‑default style, clear the style
                        if (cell.Value == null && !cell.GetStyle().Equals(workbook.DefaultStyle))
                        {
                            // Apply a default (empty) style
                            cell.SetStyle(workbook.CreateStyle());
                        }
                    }
                }

                // Remove any styles that are no longer used in the workbook
                workbook.RemoveUnusedStyles();

                // Count of instantiated Cell objects after cleanup
                long countAfter = cells.CountLarge;
                Console.WriteLine($"Instantiated cells after removing formatting‑only cells: {countAfter}");

                // Show the reduction
                long reduction = countBefore - countAfter;
                Console.WriteLine($"Reduction in instantiated cells: {reduction}");

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("FormattingOnlyCellsReductionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            FormattingOnlyCellsReductionDemo.Run();
        }
    }
}
