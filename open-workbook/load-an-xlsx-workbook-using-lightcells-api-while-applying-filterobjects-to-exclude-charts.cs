using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Create a LoadFilter that loads everything except charts
        LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        LoadFilter loadFilter = new LoadFilter(filterOptions);

        // Create a simple LightCellsDataHandler implementation
        LightCellsDataHandler handler = new SimpleHandler();

        // Configure LoadOptions with the filter and LightCells handler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = loadFilter;
        loadOptions.LightCellsDataHandler = handler;
        // Unparsed data is not needed for this read‑only scenario
        loadOptions.KeepUnparsedData = false;

        // Load the workbook using LightCells mode and the specified filter
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Verify that worksheets are loaded and charts are excluded
        Console.WriteLine("Worksheets loaded: " + workbook.Worksheets.Count);
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Sheet '{sheet.Name}' contains {sheet.Charts.Count} chart(s).");
        }
    }

    // Minimal LightCellsDataHandler that processes all sheets, rows, and cells
    class SimpleHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet)
        {
            // Process every worksheet
            return true;
        }

        public bool StartRow(int rowIndex)
        {
            // Process every row
            return true;
        }

        public bool ProcessRow(Row row)
        {
            // No custom row processing needed
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            // Process every cell in the row
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            // No custom cell processing needed
            return true;
        }
    }
}