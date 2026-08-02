using System;
using Aspose.Cells;

namespace AsposeCellsCalculationMonitorDemo
{
    // Custom monitor that inspects each cell before it is calculated
    public class CellInspectionMonitor : AbstractCalculationMonitor
    {
        private readonly Workbook _workbook;

        // Pass the workbook reference so we can access the cell being processed
        public CellInspectionMonitor(Workbook workbook)
        {
            _workbook = workbook;
        }

        // This method is called by the calculation engine before a cell is evaluated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Retrieve the cell object using the provided indexes
            Cell cell = _workbook.Worksheets[sheetIndex].Cells[rowIndex, colIndex];

            // Output information about the cell. OriginalValue is the value before calculation.
            Console.WriteLine($"Before calculating cell {cell.Name} (Sheet {sheetIndex}, Row {rowIndex}, Column {colIndex})");
            Console.WriteLine($"  Original Value: {OriginalValue}");
        }

        // Optional: you can also override AfterCalculate if you need post‑calculation info
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            Console.WriteLine($"After calculating cell {_workbook.Worksheets[sheetIndex].Cells[rowIndex, colIndex].Name}");
            Console.WriteLine($"  Value Changed: {ValueChanged}, New Value: {CalculatedValue}");
        }

        // Optional: handle circular references
        public override bool OnCircular(System.Collections.IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected during calculation.");
            return true; // Continue calculation
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data and formulas
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].Formula = "=A1*2";
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Instantiate the custom monitor, passing the workbook reference
            CellInspectionMonitor monitor = new CellInspectionMonitor(workbook);

            // Set calculation options to use the monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = monitor,
                IgnoreError = false,
                Recursive = true
            };

            // Perform formula calculation; the monitor will be invoked for each cell
            workbook.CalculateFormula(options);

            // Save the workbook to verify results
            workbook.Save("InspectionDemo.xlsx");
        }
    }
}