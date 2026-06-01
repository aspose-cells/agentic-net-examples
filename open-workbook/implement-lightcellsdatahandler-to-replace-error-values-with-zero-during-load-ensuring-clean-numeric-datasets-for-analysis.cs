using System;
using Aspose.Cells;

namespace LightCellsErrorHandlerDemo
{
    // Custom handler that replaces any error cell value with zero during loading
    public class ReplaceErrorWithZeroHandler : LightCellsDataHandler
    {
        // Process each worksheet – return true to process all sheets
        public bool StartSheet(Worksheet sheet) => true;

        // Process each row – return true to process all rows
        public bool StartRow(int rowIndex) => true;

        // Process each cell – return true to keep the cell in the workbook model after processing
        public bool StartCell(int columnIndex) => true;

        // Row processing (not needed for this task, just continue)
        public bool ProcessRow(Row row) => true;

        // Core logic: replace error values with zero
        public bool ProcessCell(Cell cell)
        {
            // If the cell contains an error, replace it with numeric zero
            if (cell.Type == CellValueType.IsError)
            {
                cell.PutValue(0);
            }
            // Keep the cell in the model
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (could be any format supported by Aspose.Cells)
            string sourcePath = "input.xlsx";

            // Configure load options to use the custom LightCellsDataHandler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new ReplaceErrorWithZeroHandler();

            // Load the workbook using the specified options (lightweight mode)
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the processed workbook – all error cells are now zero
            workbook.Save("output.xlsx");
        }
    }
}