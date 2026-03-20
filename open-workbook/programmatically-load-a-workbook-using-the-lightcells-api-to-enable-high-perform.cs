using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace LightCellsDemo
{
    // Custom handler that processes worksheet data in a streaming (low‑memory) manner
    public class MyLightCellsHandler : LightCellsDataHandler
    {
        // Called when a worksheet starts processing
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Processing sheet: {sheet.Name}");
            // Return true to continue processing this sheet
            return true;
        }

        // Called before a row is read
        public bool StartRow(int rowIndex)
        {
            // Return true to read this row
            return true;
        }

        // Called after a row's properties are read
        public bool ProcessRow(Row row)
        {
            // Example: output the row index
            Console.WriteLine($"  Row {row.Index}");
            // Return true to continue processing cells in this row
            return true;
        }

        // Called before a cell in the current row is read
        public bool StartCell(int columnIndex)
        {
            // Return true to read this cell
            return true;
        }

        // Called after a cell's data is read
        public bool ProcessCell(Cell cell)
        {
            // Output cell address and its value (if any)
            Console.WriteLine($"    Cell {cell.Name}: {cell.StringValue}");
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the large Excel file to be processed
            string inputPath = "LargeFile.xlsx";

            // Configure load options to use LightCells mode with the custom handler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new MyLightCellsHandler();

            // Load the workbook using LightCells (low‑memory) API
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the processed workbook (can be the same or a new file)
            string outputPath = "ProcessedLargeFile.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
    }
}