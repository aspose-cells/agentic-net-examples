using System;
using Aspose.Cells;

namespace LightCellsExample
{
    // Custom handler that processes cells while reading the workbook in LightCells mode
    public class LightCellsDataHandlerDemo : LightCellsDataHandler
    {
        // Called when a worksheet is about to be processed
        public override bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Starting sheet: {sheet.Name}");
            return true; // Continue processing this sheet
        }

        public override bool EndSheet(Worksheet sheet)
        {
            return true;
        }

        // Called for each row index; return true to process the row
        public override bool StartRow(int rowIndex)
        {
            Console.WriteLine($"Starting row: {rowIndex}");
            return true;
        }

        public override bool EndRow(int rowIndex)
        {
            return true;
        }

        // Called after a row is created; return true to process its cells
        public override bool ProcessRow(Row row)
        {
            Console.WriteLine($"Processing row: {row.Index}");
            return true;
        }

        // Called for each column index in the current row; return true to process the cell
        public override bool StartCell(int columnIndex)
        {
            Console.WriteLine($"Starting cell in column: {columnIndex}");
            return true;
        }

        public override bool EndCell(int columnIndex)
        {
            return true;
        }

        // Called for each cell that should be processed
        public override bool ProcessCell(Cell cell)
        {
            Console.WriteLine($"Cell[{cell.Row},{cell.Column}] = {cell.StringValue}");
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the large Excel file to be read
            string inputPath = "LargeFile_original.xlsx";

            // Create LoadOptions and assign the custom LightCellsDataHandler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new LightCellsDataHandlerDemo();

            // Load the workbook using LightCells mode (streaming, low memory)
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook after processing (demonstrates that it was loaded successfully)
            workbook.Save("ProcessedLargeFile.xlsx");
        }
    }
}