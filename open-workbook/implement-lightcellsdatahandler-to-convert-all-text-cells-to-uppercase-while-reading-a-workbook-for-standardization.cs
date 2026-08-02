// Title: Uppercase all text cells with a custom LightCellsDataHandler while loading a workbook – Aspose.Cells for .NET
// Description: Demonstrates how to subclass LightCellsDataHandler in C# to detect string cells, replace their values with an invariant uppercase version, and apply the handler via LoadOptions so the workbook is processed in LightCells mode and saved with all text standardized.
// Keywords: Aspose.Cells | LightCellsDataHandler | C# | .NET | uppercase cells | string cell conversion | LoadOptions | memory‑efficient processing | GitHub example | custom cell handler
// Common Searches: Aspose.Cells convert cells to uppercase during load | custom LightCellsDataHandler C# example | how to uppercase text in Excel with Aspose.Cells | process string cells on the fly using LightCells | load workbook with transformation handler Aspose
// Developer Intent: Create a LightCellsDataHandler that changes every string cell to uppercase while a workbook is being loaded.
// Use Cases: Normalize case of textual data across large spreadsheets before analysis. | Pre‑process incoming files to meet case‑sensitive validation rules. | Apply transformations in memory without fully materializing the workbook, improving performance for massive files.
// AI Prompts: Generate a LightCellsDataHandler in C# that trims whitespace and converts string cells to uppercase during workbook load. | Show how to configure LoadOptions with a custom LightCellsDataHandler to modify cell values on the fly in Aspose.Cells. | Explain how to chain multiple text transformations (trim, replace, uppercase) inside a LightCellsDataHandler for .NET.

using System;
using Aspose.Cells;

namespace UpperCaseLightCellsDemo
{
    // Custom LightCellsDataHandler that converts all string cell values to uppercase
    // Demonstrates how to subclass LightCellsDataHandler in C# to detect string cells, replace their values with an invariant uppercase version, and apply the handler via LoadOptions so the workbook is processed in LightCells mode and saved with all text standardized.
    public class UpperCaseHandler : LightCellsDataHandler
    {
        // Process every worksheet
        public bool StartSheet(Worksheet sheet)
        {
            // Continue processing this sheet
            return true;
        }

        // Process every row
        public bool StartRow(int rowIndex)
        {
            // Continue processing this row
            return true;
        }

        // Called after a row's properties are read; we simply continue
        public bool ProcessRow(Row row)
        {
            return true;
        }

        // Process every cell in the row
        public bool StartCell(int columnIndex)
        {
            // Continue processing this cell
            return true;
        }

        // Convert string cells to uppercase
        public bool ProcessCell(Cell cell)
        {
            // Check if the cell contains a string value
            if (cell.Type == CellValueType.IsString && !string.IsNullOrEmpty(cell.StringValue))
            {
                // Replace the cell value with its uppercase representation
                cell.PutValue(cell.StringValue.ToUpperInvariant());
            }

            // Return true to keep the cell in the workbook model
            return true;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Path for the processed workbook
            string outputPath = "output.xlsx";

            // Configure load options to use the custom LightCellsDataHandler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new UpperCaseHandler();

            // Load the workbook in LightCells mode; the handler will process each cell
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook after all text cells have been converted to uppercase
            workbook.Save(outputPath);
        }
    }
}
