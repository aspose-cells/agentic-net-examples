// Title: Set Dynamic Freeze Panes in Aspose.Cells for .NET Using Environment Variables
// Description: Creates a workbook, populates a 20×10 sample grid, reads the FREEZE_ROWS and FREEZE_COLUMNS environment variables, validates the values, and calls Worksheet.FreezePanes to freeze the specified rows and columns before saving as DynamicFreezeDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | FreezePanes | environment variables | dynamic freeze rows | dynamic freeze columns | runtime configuration | Excel export | worksheet freeze panes
// Common Searches: Aspose.Cells freeze panes from environment variable | C# set freeze rows and columns at runtime | How to use FreezePanes with variables in .NET | Dynamic freeze panes Aspose.Cells example | Read FREEZE_ROWS FREEZE_COLUMNS in C# Excel
// Developer Intent: Read FREEZE_ROWS and FREEZE_COLUMNS environment variables and apply them to Worksheet.FreezePanes so the worksheet freezes the requested rows and/or columns.
// Use Cases: Generate Excel reports where the number of frozen header rows or columns is controlled by deployment‑time environment settings. | Adjust worksheet freeze panes in a CI/CD pipeline without changing source code. | Support multi‑tenant exports where each tenant’s preferred frozen rows/columns are supplied via environment variables.
// AI Prompts: Show how to extend the sample to accept freeze settings from command‑line arguments instead of environment variables. | Provide code that validates freeze values against the worksheet size and logs a warning for out‑of‑range inputs. | Explain how to read freeze configuration from a JSON file and apply it with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicFreezeDemo
{
    // Creates a workbook, populates a 20×10 sample grid, reads the FREEZE_ROWS and FREEZE_COLUMNS environment variables, validates the values, and calls Worksheet.FreezePanes to freeze the specified rows and columns before saving as DynamicFreezeDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data so the effect of freezing can be seen
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Read environment variables that specify how many rows and columns to freeze
            // If the variables are not set or cannot be parsed, default to 0 (no freeze)
            int freezeRows = 0;
            int freezeColumns = 0;

            string rowsEnv = Environment.GetEnvironmentVariable("FREEZE_ROWS");
            string colsEnv = Environment.GetEnvironmentVariable("FREEZE_COLUMNS");

            if (!string.IsNullOrWhiteSpace(rowsEnv) && int.TryParse(rowsEnv, out int parsedRows) && parsedRows > 0)
                freezeRows = parsedRows;

            if (!string.IsNullOrWhiteSpace(colsEnv) && int.TryParse(colsEnv, out int parsedCols) && parsedCols > 0)
                freezeColumns = parsedCols;

            // Apply freeze panes only when at least one dimension is greater than zero
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // The first two parameters define the cell where the split occurs.
            // Using the same values for row/column and freezedRows/freezeColumns creates the desired freeze.
            if (freezeRows > 0 || freezeColumns > 0)
            {
                sheet.FreezePanes(freezeRows, freezeColumns, freezeRows, freezeColumns);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DynamicFreezeDemo.xlsx");
        }
    }
}
