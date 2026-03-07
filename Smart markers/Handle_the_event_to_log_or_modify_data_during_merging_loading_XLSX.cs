using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeLoadDemo
{
    // Custom handler to process cells while loading in LightCells mode
    public class LoggingAndModifyingHandler : LightCellsDataHandler
    {
        // Called before processing a worksheet
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"[StartSheet] Processing sheet: {sheet.Name}");
            return true; // Process all sheets
        }

        // Called before processing a row
        public bool StartRow(int rowIndex)
        {
            // Optionally filter rows here
            return true; // Process all rows
        }

        // Called before processing a cell in the current row
        public bool StartCell(int columnIndex)
        {
            // Process every cell
            return true;
        }

        // Called after a row's properties are read, before its cells are read
        public bool ProcessRow(Row row)
        {
            // Return true to continue processing cells in this row
            return true;
        }

        // Called for each cell that is read
        public bool ProcessCell(Cell cell)
        {
            // Log basic information
            Console.WriteLine($"[ProcessCell] Cell {cell.Name} (R{cell.Row}, C{cell.Column}) - Value: {cell.Value}");

            // Log if the cell is part of a merged range
            if (cell.IsMerged)
            {
                Console.WriteLine($"    -> Cell {cell.Name} is merged.");
            }

            // Example modification: double numeric values
            if (cell.IsNumericValue)
            {
                double original = cell.DoubleValue;
                double modified = original * 2;
                cell.PutValue(modified);
                Console.WriteLine($"    -> Numeric value changed from {original} to {modified}");
            }

            // Example modification: prepend text to string values
            if (cell.Type == CellValueType.IsString)
            {
                string original = cell.StringValue;
                string modified = $"[Modified] {original}";
                cell.PutValue(modified);
                Console.WriteLine($"    -> String value changed to \"{modified}\"");
            }

            // Return true to keep the cell in the in‑memory model (optional)
            // Returning false reduces memory usage; here we keep it for possible later use
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (replace with your actual file)
            string sourcePath = "InputWorkbook.xlsx";

            // Ensure the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Create load options and attach the custom LightCellsDataHandler
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.LightCellsDataHandler = new LoggingAndModifyingHandler();

            // Load the workbook using LightCells mode (streaming)
            Workbook workbook = new Workbook(sourcePath, loadOptions);
            Console.WriteLine("Workbook loaded with custom LightCellsDataHandler.");

            // Optional: set save options to merge conditional formatting/validation areas
            XlsSaveOptions saveOptions = new XlsSaveOptions
            {
                MergeAreas = true,               // Merge conditional formatting and validation areas
                ValidateMergedAreas = true       // Validate merged cells before saving
            };

            // Save the processed workbook
            string outputPath = "ProcessedWorkbook.xlsx";
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Processed workbook saved to: {outputPath}");
        }
    }
}