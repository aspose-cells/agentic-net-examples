using System;
using Aspose.Cells;

namespace LightCellsMonitoringExample
{
    class Program
    {
        static void Main()
        {
            // Create an instance of the custom handler that will receive callbacks during loading.
            var handler = new MonitoringHandler();

            // Configure load options to use the handler.
            var loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = handler;

            // Load the workbook with the specified options.
            // The handler's methods will be invoked for each sheet, row, and cell.
            var workbook = new Workbook("input.xlsx", loadOptions);

            // After loading you can work with the workbook as usual.
            Console.WriteLine($"Workbook loaded. Total worksheets: {workbook.Worksheets.Count}");

            // Save the workbook (no changes made, just demonstrating the save lifecycle).
            workbook.Save("output.xlsx");
        }

        // Custom implementation of LightCellsDataHandler to monitor loading stages.
        class MonitoringHandler : LightCellsDataHandler
        {
            // Called before any cells of a worksheet are processed.
            public bool StartSheet(Worksheet sheet)
            {
                Console.WriteLine($"StartSheet: Processing worksheet \"{sheet.Name}\" (Index {sheet.Index})");
                // Return true to continue processing this sheet.
                return true;
            }

            // Called before a row is processed.
            public bool StartRow(int rowIndex)
            {
                Console.WriteLine($"StartRow: Row {rowIndex}");
                // Return true to continue processing this row.
                return true;
            }

            // Called after a row's metadata is read; you can inspect the Row object here.
            public bool ProcessRow(Row row)
            {
                Console.WriteLine($"ProcessRow: Row index {row.Index}");
                // Return true to continue processing cells in this row.
                return true;
            }

            // Called before a cell in the current row is processed.
            public bool StartCell(int columnIndex)
            {
                Console.WriteLine($"StartCell: Column {columnIndex}");
                // Return true to process this cell.
                return true;
            }

            // Called for each cell that is to be processed.
            public bool ProcessCell(Cell cell)
            {
                // Display cell address and its string representation.
                Console.WriteLine($"ProcessCell: {cell.Name} = \"{cell.StringValue}\"");
                // Return true to continue processing subsequent cells.
                return true;
            }
        }
    }
}