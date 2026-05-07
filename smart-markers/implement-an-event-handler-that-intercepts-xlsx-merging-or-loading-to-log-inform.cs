using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a custom LightCellsDataHandler to log and optionally modify cell data during loading
        var handler = new MergeLoggingHandler();

        // Configure LoadOptions to use the custom handler
        var loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.LightCellsDataHandler = handler;

        // Load the workbook with the specified options
        using (var workbook = new Workbook("input.xlsx", loadOptions))
        {
            // After loading, enumerate merged areas to demonstrate that they are available
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (CellArea area in sheet.Cells.MergedCells)
                {
                    Console.WriteLine($"Merged area in sheet '{sheet.Name}': " +
                                      $"Rows {area.StartRow}-{area.EndRow}, " +
                                      $"Columns {area.StartColumn}-{area.EndColumn}");
                }
            }

            // Prepare SaveOptions to validate merged cells before saving
            var saveOptions = new XlsbSaveOptions
            {
                ValidateMergedAreas = true, // ensure merged cells are consistent
                MergeAreas = true           // optionally merge conditional formatting/validation areas
            };

            // Save the processed workbook
            workbook.Save("output.xlsb", saveOptions);
        }
    }

    // Custom LightCellsDataHandler implementation
    class MergeLoggingHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Start processing sheet: {sheet.Name}");
            return true; // continue processing this sheet
        }

        public bool StartRow(int rowIndex)
        {
            Console.WriteLine($"  Starting row: {rowIndex}");
            return true; // continue processing this row
        }

        public bool ProcessRow(Row row)
        {
            // No row-level modification needed; keep processing
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            Console.WriteLine($"    Starting cell at column: {columnIndex}");
            return true; // continue processing this cell
        }

        public bool ProcessCell(Cell cell)
        {
            // Log cell coordinates and its current value
            Console.WriteLine($"      Cell[{cell.Row},{cell.Column}] Value: {cell.Value}");

            // Example modification: prepend a marker to string cells
            if (cell.Type == CellValueType.IsString)
            {
                cell.PutValue("LOG:" + cell.StringValue);
            }

            // Return true to keep the cell in the workbook model after processing
            return true;
        }
    }
}