using System;
using Aspose.Cells;

namespace LightCellsFirstThreeSheetsDemo
{
    // Author: Aspose.Cells .NET example – loads only the first three worksheets using LightCells
    class Program
    {
        static void Main()
        {
            // Create a custom LightCellsDataHandler that processes only the first three sheets
            var handler = new FirstThreeSheetsHandler();

            // Configure LoadOptions to use the handler
            var loadOptions = new LoadOptions
            {
                LightCellsDataHandler = handler
            };

            // Load the workbook with LightCells mode
            var workbook = new Workbook("sample.xlsx", loadOptions);

            // Optional: remove any remaining worksheets to free memory completely
            while (workbook.Worksheets.Count > 3)
            {
                // Remove from the end to avoid shifting indices
                workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);
            }

            // Demonstrate that only three worksheets are present
            Console.WriteLine("Worksheets loaded: " + workbook.Worksheets.Count);
        }
    }

    // Custom LightCellsDataHandler that allows processing of only the first three worksheets
    class FirstThreeSheetsHandler : LightCellsDataHandler
    {
        private int _processedSheets = 0;

        // Called when a worksheet starts processing; return true to load it, false to skip
        public bool StartSheet(Worksheet sheet)
        {
            if (_processedSheets < 3)
            {
                _processedSheets++;
                return true; // load this sheet
            }
            return false; // skip further sheets
        }

        // Continue processing rows
        public bool StartRow(int rowIndex) => true;
        public bool ProcessRow(Row row) => true;

        // Continue processing cells
        public bool StartCell(int columnIndex) => true;
        public bool ProcessCell(Cell cell) => true;

        // No need to gather strings for this demo
        public bool IsGatherString() => false;
    }
}