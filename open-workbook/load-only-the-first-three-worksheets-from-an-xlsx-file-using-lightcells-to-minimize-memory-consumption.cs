using System;
using Aspose.Cells;

namespace LightCellsFirstThreeSheets
{
    // Custom handler that processes only the first three worksheets
    public class FirstThreeSheetsHandler : LightCellsDataHandler
    {
        // Called before reading a worksheet; return true only for sheets 0,1,2
        public bool StartSheet(Worksheet sheet)
        {
            // Process only if the sheet index is less than 3
            return sheet.Index < 3;
        }

        // Called before reading each row; we want to read all rows in the selected sheets
        public bool StartRow(int rowIndex) => true;

        // Called after a row is read; no special processing needed
        public bool ProcessRow(Row row) => true;

        // Called before reading each cell; we want to read all cells in the selected rows
        public bool StartCell(int columnIndex) => true;

        // Called after a cell is read; no special processing needed
        public bool ProcessCell(Cell cell) => true;
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourceFile = "LargeWorkbook.xlsx";

            // Create LoadOptions and assign the custom LightCellsDataHandler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new FirstThreeSheetsHandler();

            // Load the workbook using the LightCells mode (only first three sheets will be loaded)
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Verify that only three worksheets are present
            Console.WriteLine($"Worksheets loaded: {workbook.Worksheets.Count}");
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Console.WriteLine($"Sheet {i}: {workbook.Worksheets[i].Name}");
            }

            // Optionally save the partially loaded workbook to a new file
            string outputFile = "FirstThreeSheets.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputFile}");
        }
    }
}