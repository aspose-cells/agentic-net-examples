using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook using a custom LightCellsDataHandler that upper‑cases text cells
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new UpperCaseHandler();

        // The handler will be invoked during loading
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the processed workbook
        workbook.Save("output.xlsx");
    }

    // Custom handler: converts every string cell to upper case while reading
    private class UpperCaseHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet) => true;          // Process all sheets
        public bool StartRow(int rowIndex) => true;              // Process all rows
        public bool StartCell(int columnIndex) => true;          // Process all cells
        public bool ProcessRow(Row row) => true;                 // No row‑level changes

        public bool ProcessCell(Cell cell)
        {
            // If the cell contains a string, replace it with its upper‑case version
            if (cell.Type == CellValueType.IsString)
            {
                string upper = cell.StringValue?.ToUpperInvariant();
                cell.PutValue(upper);
            }
            return true; // Continue processing
        }
    }
}

// Author: Aspose.Cells .NET example – custom LightCellsDataHandler for text standardization.