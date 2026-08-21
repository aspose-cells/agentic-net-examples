// Title: Aspose.Cells .NET – Replace error cells with zero using LightCellsDataHandler on load
// Description: Demonstrates a custom LightCellsDataHandler (ZeroErrorHandler) that intercepts each cell during workbook loading, detects error values, and substitutes them with numeric zero. The handler is attached via LoadOptions, enabling clean numeric data for downstream analysis.
// Keywords: Aspose.Cells | .NET | LightCellsDataHandler | replace error cells | set error to zero | Excel load options | data cleaning | error handling in Excel | large workbook processing | performance‑optimized import
// Common Searches: Aspose.Cells LightCellsDataHandler replace error values | set #N/A cells to 0 during workbook load .NET | custom cell handler for error handling Aspose | load Excel file and convert errors to zero | how to clean error cells with Aspose.Cells
// Developer Intent: Create a LightCellsDataHandler that converts any error‑type cell to numeric zero while loading a workbook.
// Use Cases: Prepare raw Excel data for statistical analysis by eliminating #DIV/0! and #N/A errors. | Prevent runtime calculation failures in downstream .NET processing pipelines. | Standardize import routines for financial or scientific datasets that contain sporadic error cells.
// AI Prompts: Generate a LightCellsDataHandler that logs the address of each error cell before setting it to zero. | Extend ZeroErrorHandler to also replace empty string cells with null values during load. | Show how to combine the custom handler with LoadOptions for high‑performance processing of a 500 MB workbook containing thousands of error cells.

using System;
using Aspose.Cells;

namespace AsposeCellsLightCellsHandlerDemo
{
    // Custom handler that replaces any error cell value with zero during loading.
    // Demonstrates a custom LightCellsDataHandler (ZeroErrorHandler) that intercepts each cell during workbook loading, detects error values, and substitutes them with numeric zero. The handler is attached via LoadOptions, enabling clean numeric data for downstream analysis.
    public class ZeroErrorHandler : LightCellsDataHandler
    {
        // Process each worksheet – return true to process all sheets.
        public bool StartSheet(Worksheet sheet)
        {
            // Optionally, you can log the sheet name here.
            return true;
        }

        // Process each row – return true to process all rows.
        public bool StartRow(int rowIndex)
        {
            return true;
        }

        // Decide whether to process a particular cell – return true for all cells.
        public bool StartCell(int columnIndex)
        {
            return true;
        }

        // Called after a cell's data has been read.
        public bool ProcessCell(Cell cell)
        {
            // If the cell contains an error value, replace it with numeric zero.
            if (cell.Type == CellValueType.IsError)
            {
                cell.PutValue(0);
            }

            // Return true to keep the cell in the workbook model.
            return true;
        }

        // Called after a row's data has been read – we don't need special handling.
        public bool ProcessRow(Row row)
        {
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be any supported format).
            string sourcePath = "input.xlsx";

            // Path where the processed workbook will be saved.
            string destinationPath = "output.xlsx";

            // Create load options and assign the custom LightCellsDataHandler.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new ZeroErrorHandler();

            // Load the workbook using the specified load options.
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook – all error cells are now zero.
            workbook.Save(destinationPath);
        }
    }
}
