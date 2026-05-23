using System;
using Aspose.Cells;
using System.Collections;

namespace AsposeCellsProgressDemo
{
    // Custom monitor that reports calculation progress as a percentage
    public class ProgressCalculationMonitor : AbstractCalculationMonitor
    {
        private readonly int _totalCells;      // Total number of cells that will be calculated
        private int _processedCells;           // Number of cells already processed

        public ProgressCalculationMonitor(int totalCells)
        {
            _totalCells = totalCells;
            _processedCells = 0;
        }

        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            _processedCells++;
            double percent = (double)_processedCells / _totalCells * 100.0;
            Console.WriteLine($"Calculating cell [{sheetIndex}, {rowIndex}, {columnIndex}] - {percent:F2}% completed");
        }

        // Optional: after calculation you can also output details
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Example: show if the value changed
            if (ValueChanged)
            {
                Console.WriteLine($"Cell [{sheetIndex}, {rowIndex}, {columnIndex}] value changed from {OriginalValue} to {CalculatedValue}");
            }
        }

        // Handle circular references (just use default behavior)
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            return base.OnCircular(circularCellsData);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large number of formulas to simulate an extensive calculation
            int rows = 1000;
            for (int i = 0; i < rows; i++)
            {
                // Simple cumulative sum formula
                cells[$"A{i + 1}"].PutValue(i + 1);
                cells[$"B{i + 1}"].Formula = $"=SUM(A1:A{i + 1})";
            }

            // Determine total number of formula cells that will be processed
            int totalFormulaCells = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Cell cell in ws.Cells)
                {
                    if (cell.IsFormula)
                        totalFormulaCells++;
                }
            }

            // Set up calculation options with the custom progress monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new ProgressCalculationMonitor(totalFormulaCells)
            };

            // Perform the calculation with progress reporting
            workbook.CalculateFormula(options);

            // Save the workbook (using the standard save lifecycle)
            workbook.Save("ProgressCalculationDemo.xlsx");
        }
    }
}