using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AdvancedLoadingDemo
{
    // Custom LightCellsDataHandler to process cells while loading
    public class CustomLightCellsDataHandler : LightCellsDataHandler
    {
        // Called when a sheet starts loading
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Start processing sheet: {sheet.Name}");
            return true; // Continue processing this sheet
        }

        // Called when a row starts loading
        public bool StartRow(int rowIndex)
        {
            // Example: skip processing rows beyond a certain index for performance
            if (rowIndex > 1000)
                return false; // Skip remaining rows in this sheet
            return true;
        }

        // Called for each cell being loaded
        public bool ProcessCell(Cell cell)
        {
            // Example: log only numeric cells
            if (cell.Type == CellValueType.IsNumeric)
                Console.WriteLine($"Numeric cell [{cell.Row}, {cell.Column}] = {cell.Value}");
            return true; // Continue processing
        }

        // Required overrides (not used in this demo)
        public bool StartCell(int columnIndex) => true;
        public bool ProcessRow(Row row) => true;
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "sample.xlsx";

            // Configure advanced load options
            LoadOptions loadOptions = new LoadOptions
            {
                // Do not keep unparsed data to reduce memory usage
                KeepUnparsedData = false,

                // Prefer memory-efficient loading
                MemorySetting = MemorySetting.MemoryPreference,

                // Skip formula parsing on open for faster load
                ParsingFormulaOnOpen = false,

                // Use a custom LightCellsDataHandler for row/column filtering
                LightCellsDataHandler = new CustomLightCellsDataHandler(),

                // Load only cell data and charts (skip pictures, shapes, etc.)
                LoadFilter = new LoadFilter(LoadDataFilterOptions.CellData | LoadDataFilterOptions.Chart)
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access the first worksheet to demonstrate that data is available
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"Loaded worksheet: {sheet.Name}");
            Console.WriteLine($"Total rows: {sheet.Cells.MaxDataRow + 1}");
            Console.WriteLine($"Total columns: {sheet.Cells.MaxDataColumn + 1}");

            // Example: read a specific cell if it exists
            if (sheet.Cells["A1"].Value != null)
                Console.WriteLine($"A1 = {sheet.Cells["A1"].Value}");

            // Save the workbook after loading (no modifications in this demo)
            string outputPath = "sample_loaded.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputPath}");

            // Dispose the workbook to release resources
            workbook.Dispose();
        }
    }
}