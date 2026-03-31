using System;
using Aspose.Cells;

class LightCellsReadDemo
{
    static void Main()
    {
        // Create a custom LightCellsDataHandler to process cells while reading
        var handler = new MyLightCellsDataHandler();

        // Set up load options and assign the handler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = handler;

        // Initialize the workbook with LightCells mode to read a large XLSX file
        Workbook workbook = new Workbook("LargeFile.xlsx", loadOptions);

        // Optionally save the workbook after processing
        workbook.Save("ProcessedLargeFile.xlsx");
    }

    // Implementation of LightCellsDataHandler
    class MyLightCellsDataHandler : LightCellsDataHandler
    {
        // Called when a worksheet starts processing
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Starting sheet: {sheet.Name}");
            return true; // Continue processing this sheet
        }

        // Called before a row is processed
        public bool StartRow(int rowIndex)
        {
            // Return true to process the row
            return true;
        }

        // Called after a row's properties are read
        public bool ProcessRow(Row row)
        {
            Console.WriteLine($"Processing row: {row.Index}");
            return true; // Continue to cells in this row
        }

        // Called before a cell is processed
        public bool StartCell(int columnIndex)
        {
            // Return true to process the cell
            return true;
        }

        // Called for each cell that is processed
        public bool ProcessCell(Cell cell)
        {
            Console.WriteLine($"Cell {cell.Name} = {cell.Value}");
            return true; // Continue processing
        }
    }
}