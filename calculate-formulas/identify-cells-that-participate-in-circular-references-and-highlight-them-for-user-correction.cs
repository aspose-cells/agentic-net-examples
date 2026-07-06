using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace CircularReferenceHighlighter
{
    // Custom monitor to capture circular reference cells
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        private readonly Worksheet _worksheet;
        private readonly List<Cell> _circularCells = new List<Cell>();

        public CircularReferenceMonitor(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Collect all cells involved in the circular reference
            while (circularCellsData.MoveNext())
            {
                // The enumerated object is a CalculationCell; its Cell property gives the actual Cell
                // In many examples the object can be cast directly to Cell, so we handle both cases
                var obj = circularCellsData.Current;
                Cell cell = null;

                // Try to get Cell from CalculationCell if possible
                var type = obj.GetType();
                var cellProp = type.GetProperty("Cell");
                if (cellProp != null)
                {
                    cell = cellProp.GetValue(obj) as Cell;
                }
                else
                {
                    cell = obj as Cell;
                }

                if (cell != null)
                {
                    _circularCells.Add(cell);
                }
            }

            // Return true to let the engine continue calculation (or false to stop)
            return true;
        }

        // After calculation finishes, highlight the collected cells
        public void HighlightCircularCells()
        {
            foreach (var cell in _circularCells)
            {
                // Create a style with a yellow background
                Style style = cell.GetStyle();
                style.ForegroundColor = Color.Yellow;
                style.Pattern = BackgroundType.Solid;
                cell.SetStyle(style);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a circular reference scenario
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Optional: add more data to demonstrate normal cells
            sheet.Cells["C1"].PutValue(10);
            sheet.Cells["D1"].Formula = "=C1*2";

            // Instantiate the custom monitor, passing the worksheet for later styling
            var monitor = new CircularReferenceMonitor(sheet);

            // Configure calculation options to use the monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = monitor,
                Recursive = true // default, but explicit for clarity
            };

            // Perform formula calculation; circular references will be captured by the monitor
            workbook.CalculateFormula(options);

            // Highlight cells that participated in circular references
            monitor.HighlightCircularCells();

            // Save the workbook (the highlighted cells will be visible in the saved file)
            workbook.Save("CircularReferenceHighlighted.xlsx");
        }
    }
}