// Title: How to Detect and Highlight Circular Reference Cells in an Excel Workbook Using Aspose.Cells for .NET
// AI Prompts: Write C# code that implements a custom CalculationMonitor to collect cells participating in circular references and applies a yellow background style to each of those cells. | Demonstrate configuring CalculationOptions with Recursive = true and a CircularReferenceMonitor, then running workbook.CalculateFormula and saving the workbook with the highlighted circular cells. | Adapt the example to log the addresses of circular reference cells to a text file while still applying a highlight style to them in the worksheet.
// Common Searches: Aspose.Cells .NET highlight cells that cause circular reference errors | C# example using AbstractCalculationMonitor to find circular formulas in Excel | How to apply a style to cells detected by a custom calculation monitor in Aspose.Cells | Detect self‑referencing formulas and mark them in an Excel file with Aspose.Cells
// Tags: Aspose.Cells circular reference detection | custom CalculationMonitor C# | highlight circular cells Excel | apply style to error cells Aspose | calculate formulas with monitor Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceHighlight
{
    // Monitor to capture cells involved in circular references during calculation
    // The sample creates a workbook with circular formulas, defines a CircularReferenceMonitor derived from AbstractCalculationMonitor to capture each cell involved in a circular reference during calculation, runs workbook.CalculateFormula with this monitor, creates a yellow solid background style, applies the style to all captured cells, and saves the highlighted workbook as CircularReferenceHighlighted.xlsx.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        private readonly Workbook _workbook;
        public List<Cell> CircularCells { get; } = new List<Cell>();

        public CircularReferenceMonitor(Workbook workbook)
        {
            _workbook = workbook;
        }

        // Called by the calculation engine when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            try
            {
                while (circularCellsData.MoveNext())
                {
                    var current = circularCellsData.Current;
                    if (current != null)
                    {
                        // Use dynamic to access SheetIndex, Row, and Column properties
                        dynamic cellInfo = current;
                        int sheetIdx = cellInfo.SheetIndex;
                        int row = cellInfo.Row;
                        int col = cellInfo.Column;

                        Cell cell = _workbook.Worksheets[sheetIdx].Cells[row, col];
                        CircularCells.Add(cell);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing circular reference data: {ex.Message}");
            }

            // Continue calculation
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ------------------- Create Workbook -------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set up a circular reference scenario
                cells["A1"].Formula = "=B1";
                cells["B1"].Formula = "=A1";
                cells["C1"].Formula = "=C1"; // self‑reference

                // ------------------- Set Calculation Options -------------------
                CircularReferenceMonitor monitor = new CircularReferenceMonitor(workbook);
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = monitor,
                    Recursive = true
                };

                // Perform calculation; monitor will collect circular cells
                workbook.CalculateFormula(options);

                // ------------------- Highlight Circular Cells -------------------
                // Define a style with a distinct background color
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;

                foreach (Cell circularCell in monitor.CircularCells)
                {
                    circularCell.SetStyle(highlightStyle);
                }

                // ------------------- Save Workbook -------------------
                string outputPath = "CircularReferenceHighlighted.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
