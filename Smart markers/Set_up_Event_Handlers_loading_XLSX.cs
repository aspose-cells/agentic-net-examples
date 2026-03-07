using System;
using Aspose.Cells;

namespace AsposeCellsEventHandlerDemo
{
    // Custom handler that processes cells while the workbook is being loaded in LightCells mode
    public class CustomLightCellsDataHandler : LightCellsDataHandler
    {
        // Called before a worksheet is processed
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Start processing sheet: {sheet.Name}");
            return true; // Continue processing this sheet
        }

        // Called before a row is processed
        public bool StartRow(int rowIndex)
        {
            Console.WriteLine($"  Start row: {rowIndex}");
            return true; // Continue processing this row
        }

        // Called after a row object is created; can be used to read row properties
        public bool ProcessRow(Row row)
        {
            // No special processing needed; just continue
            return true;
        }

        // Called before a cell in the current row is processed
        public bool StartCell(int columnIndex)
        {
            Console.WriteLine($"    Start cell at column: {columnIndex}");
            return true; // Continue processing this cell
        }

        // Called for each cell that is to be processed
        public bool ProcessCell(Cell cell)
        {
            Console.WriteLine($"      Cell[{cell.Row},{cell.Column}] Value: {cell.Value}");
            return true; // Continue processing subsequent cells
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare load options
            LoadOptions loadOptions = new LoadOptions();

            // Attach the custom LightCellsDataHandler
            loadOptions.LightCellsDataHandler = new CustomLightCellsDataHandler();

            // Example: skip formula parsing to speed up loading (optional)
            loadOptions.ParsingFormulaOnOpen = false;

            // Load the XLSX file using the options above
            // The constructor Workbook(string, LoadOptions) follows the provided rule
            Workbook workbook = new Workbook("sample.xlsx", loadOptions);

            // After loading, you can work with the workbook as usual
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine($"First worksheet name: {firstSheet.Name}");
            Console.WriteLine($"Cell A1 value after load: {firstSheet.Cells["A1"].Value}");

            // Save the workbook to demonstrate that it was loaded successfully
            // The Save(string) method follows the provided rule
            workbook.Save("ProcessedSample.xlsx");
        }
    }
}