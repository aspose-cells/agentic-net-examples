using System;
using System.Collections;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceHighlight
{
    // Custom monitor to capture circular reference cells and highlight them
    public class CircularReferenceHighlighter : AbstractCalculationMonitor
    {
        private readonly Workbook _workbook;

        public CircularReferenceHighlighter(Workbook workbook)
        {
            _workbook = workbook;
        }

        // Called when the calculation engine detects a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected. Highlighting involved cells...");

            while (circularCellsData.MoveNext())
            {
                try
                {
                    var item = circularCellsData.Current;
                    if (item == null) continue;

                    Cell cell = null;

                    // If the item is already a Cell instance, use it directly
                    if (item is Cell directCell)
                    {
                        cell = directCell;
                    }
                    else
                    {
                        // Otherwise try to obtain sheet, row and column via dynamic members
                        dynamic dyn = item;
                        int sheetIdx = dyn.SheetIndex;
                        int row = dyn.Row;
                        int col = dyn.Column;

                        cell = _workbook.Worksheets[sheetIdx].Cells[row, col];
                    }

                    // Apply a yellow background to indicate the problem
                    Style style = cell.GetStyle();
                    style.ForegroundColor = Color.Yellow;
                    style.Pattern = BackgroundType.Solid;
                    cell.SetStyle(style);

                    Console.WriteLine($"Highlighted cell: {cell.Name}");
                }
                catch (Exception ex)
                {
                    // Log and continue with next item
                    Console.WriteLine($"Failed to highlight a circular cell: {ex.Message}");
                }
            }

            // Continue normal calculation after highlighting
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (could also be loaded from a file)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set up a simple circular reference: A1 -> B1 -> A1
                cells["A1"].Formula = "=B1";
                cells["B1"].Formula = "=A1";

                // Prepare calculation options with our custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceHighlighter(workbook),
                    Recursive = true // keep default recursive behavior
                };

                // Perform calculation; the monitor will highlight circular cells
                workbook.CalculateFormula(options);

                // Save the workbook so the user can see highlighted cells
                string outputPath = "CircularReferenceHighlighted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with highlighted circular reference cells: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}