using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace LightCellsNumericReader
{
    // Author: Aspose.Cells example – reads numeric cells via LightCells API into a high‑performance collection
    class Program
    {
        static void Main()
        {
            // Prepare a collection to hold numeric values (high‑performance in‑memory storage)
            var numericValues = new List<double>();

            // Create a custom LightCellsDataHandler that captures numeric cell values
            var handler = new NumericCaptureHandler(numericValues);

            // Set up load options to use LightCells mode with the custom handler
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.LightCellsDataHandler = handler;

            // Load the workbook using LightCells (the handler will be invoked during loading)
            var workbook = new Workbook("input.xlsx", loadOptions);

            // At this point, numericValues contains all numeric cell values from the workbook
            Console.WriteLine($"Total numeric cells captured: {numericValues.Count}");
            foreach (var val in numericValues)
            {
                Console.WriteLine(val);
            }

            // Optionally, save the workbook (unchanged) to demonstrate lifecycle compliance
            workbook.Save("output.xlsx");
        }
    }

    // Custom LightCellsDataHandler implementation
    public class NumericCaptureHandler : LightCellsDataHandler
    {
        private readonly IList<double> _numericValues;

        public NumericCaptureHandler(IList<double> numericValues)
        {
            _numericValues = numericValues;
        }

        // Process all sheets
        public bool StartSheet(Worksheet sheet) => true;

        // Process all rows
        public bool StartRow(int rowIndex) => true;

        // Process all cells
        public bool StartCell(int columnIndex) => true;

        // Capture numeric values; return false to avoid keeping the cell in the model
        public bool ProcessCell(Cell cell)
        {
            if (cell.IsNumericValue)
            {
                // Store the numeric value in the collection
                _numericValues.Add(cell.DoubleValue);
            }
            // Do not keep the cell in the workbook model to preserve memory efficiency
            return false;
        }

        // Continue processing rows
        public bool ProcessRow(Row row) => true;
    }
}