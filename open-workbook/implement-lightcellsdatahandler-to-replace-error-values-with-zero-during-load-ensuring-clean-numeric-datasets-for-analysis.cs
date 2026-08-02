// Title: Clean Excel error values to zero with a custom LightCellsDataHandler in Aspose.Cells for .NET
// Description: Demonstrates how to attach a LightCellsDataHandler to LoadOptions, stream a workbook with Aspose.Cells, detect cells where CellValueType.IsError is true, replace each error with numeric zero using PutValue, and save the cleaned file. Ideal for fast, memory‑efficient preprocessing of large spreadsheets.
// Keywords: Aspose.Cells | LightCellsDataHandler | C# | .NET | replace Excel error cells | CellValueType.IsError | streaming workbook load | data cleaning | error to zero conversion | Excel preprocessing
// Common Searches: Aspose.Cells replace #DIV/0! with zero | LightCellsDataHandler example C# | how to remove Excel errors during load | custom handler for error cells Aspose | streaming load clean numeric data
// Developer Intent: Use a LightCellsDataHandler to convert every error cell to zero while loading an Excel workbook with Aspose.Cells.
// Use Cases: Prepare large financial or scientific datasets for statistical analysis by eliminating error values during import. | Prevent runtime exceptions in calculations that require numeric inputs by normalizing error cells to zero. | Accelerate ETL pipelines by handling error cleanup in the streaming load phase rather than in a separate pass.
// AI Prompts: Create a LightCellsDataHandler that replaces error cells with a configurable default value and logs the cell addresses. | Extend the provided handler to also convert blank cells to a user‑defined placeholder. | Write NUnit tests confirming that error cells become zero and that non‑error cells remain unchanged after processing.

using System;
using Aspose.Cells;

// Demonstrates how to attach a LightCellsDataHandler to LoadOptions, stream a workbook with Aspose.Cells, detect cells where CellValueType.IsError is true, replace each error with numeric zero using PutValue, and save the cleaned file. Ideal for fast, memory‑efficient preprocessing of large spreadsheets.
class Program
{
    static void Main()
    {
        // Input workbook that may contain error values
        string inputFile = "input.xlsx";

        // Output workbook with errors replaced by zero
        string outputFile = "output_clean.xlsx";

        // Create load options and attach the custom LightCellsDataHandler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new ErrorReplacingHandler();

        // Load the workbook in LightCells mode – the handler will process each cell
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // Save the cleaned workbook
        workbook.Save(outputFile);
    }

    // Custom LightCellsDataHandler implementation
    private class ErrorReplacingHandler : LightCellsDataHandler
    {
        // Process all sheets
        public bool StartSheet(Worksheet sheet) => true;

        // Process all rows
        public bool StartRow(int rowIndex) => true;

        // No special row processing needed
        public bool ProcessRow(Row row) => true;

        // Process all cells in a row
        public bool StartCell(int columnIndex) => true;

        // Replace error cells with zero
        public bool ProcessCell(Cell cell)
        {
            // Detect error value (e.g., #DIV/0!, #VALUE!, etc.)
            if (cell.Type == CellValueType.IsError)
            {
                // Replace the error with numeric zero
                cell.PutValue(0);
            }

            // Return true to keep the cell in the workbook model after processing
            return true;
        }
    }
}
