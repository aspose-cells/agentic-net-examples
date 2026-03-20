using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace LightCellsDemo
{
    // Event‑driven handler that processes cells while the workbook is being loaded.
    class MyLightCellsHandler : LightCellsDataHandler
    {
        // Called before a worksheet is read. Return true to process this sheet.
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Start processing sheet: {sheet.Name}");
            return true;
        }

        // Called before a row is read. Return true to process this row.
        public bool StartRow(int rowIndex)
        {
            // Example: skip header row (index 0)
            return rowIndex != 0;
        }

        // Called after a row is created. Return true to continue processing its cells.
        public bool ProcessRow(Row row)
        {
            Console.WriteLine($"Processing Row {row.Index}");
            return true;
        }

        // Called before a cell in the current row is read. Return true to process this cell.
        public bool StartCell(int columnIndex)
        {
            return true;
        }

        // Called for each cell that is read. Here we simply output its address and value.
        public bool ProcessCell(Cell cell)
        {
            Console.WriteLine($"Cell {cell.Name} = {cell.StringValue}");
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to a large Excel file that we want to read efficiently.
            string inputFile = "LargeData.xlsx";

            // Configure LoadOptions to use the LightCellsDataHandler.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new MyLightCellsHandler();

            // Load the workbook in LightCells (event‑driven) mode.
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // After processing, optionally save the workbook using normal save options.
            string outputFile = "ProcessedData.xlsx";
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
            workbook.Save(outputFile, saveOptions);

            Console.WriteLine("Workbook loaded and saved successfully.");
        }
    }
}