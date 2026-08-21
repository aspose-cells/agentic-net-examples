// Title: Uppercase All Text Cells with a Custom LightCellsDataHandler in Aspose.Cells for .NET
// Description: Demonstrates how to create an UpperCaseHandler that inherits LightCellsDataHandler, converts every string cell to uppercase during workbook loading, and saves the transformed file using Aspose.Cells LoadOptions in LightCells mode.
// Keywords: Aspose.Cells LightCellsDataHandler | C# uppercase cell values | convert string cells to uppercase | load workbook LightCells mode | custom cell processing Aspose | memory‑efficient Excel transformation | Aspose.Cells .NET example | uppercase text during import
// Common Searches: Aspose.Cells LightCells handler to uppercase text | C# convert all string cells to uppercase while loading | custom LightCellsDataHandler example | load Excel file with case conversion using Aspose | process cells on the fly Aspose.Cells
// Developer Intent: Create a LightCellsDataHandler that changes every string cell to uppercase during workbook loading.
// Use Cases: Standardize textual data by forcing uppercase when importing large Excel files without full in‑memory loading. | Apply on‑the‑fly transformations (e.g., trimming, case conversion) to reduce post‑processing steps. | Generate a cleaned copy of an input workbook for downstream analytics or reporting.
// AI Prompts: Write a LightCellsDataHandler in C# that trims whitespace and converts cell text to title case while loading a workbook. | Show how to use Aspose.Cells LoadOptions with a custom handler to replace specific substrings in string cells. | Explain how to chain multiple transformations (e.g., trim, uppercase, replace) inside the ProcessCell method of a LightCellsDataHandler.

using System;
using Aspose.Cells;

namespace UpperCaseLightCellsDemo
{
    // Custom LightCellsDataHandler that converts all string cell values to uppercase
    // Demonstrates how to create an UpperCaseHandler that inherits LightCellsDataHandler, converts every string cell to uppercase during workbook loading, and saves the transformed file using Aspose.Cells LoadOptions in LightCells mode.
    public class UpperCaseHandler : LightCellsDataHandler
    {
        // Process every worksheet
        public bool StartSheet(Worksheet sheet)
        {
            // Process all sheets
            return true;
        }

        // Process every row
        public bool StartRow(int rowIndex)
        {
            // Process all rows
            return true;
        }

        // Called after a row is read; continue processing its cells
        public bool ProcessRow(Row row)
        {
            return true;
        }

        // Prepare to process each cell; process all cells
        public bool StartCell(int columnIndex)
        {
            return true;
        }

        // Convert string values to uppercase
        public bool ProcessCell(Cell cell)
        {
            // Check if the cell contains a string value
            if (cell.Type == CellValueType.IsString && !string.IsNullOrEmpty(cell.StringValue))
            {
                // Replace the value with its uppercase representation
                cell.PutValue(cell.StringValue.ToUpper());
            }

            // Keep the cell in the workbook model
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with actual file path)
            string sourcePath = "input.xlsx";

            // Path to the processed workbook
            string outputPath = "output.xlsx";

            // Create load options and assign the custom handler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new UpperCaseHandler();

            // Load the workbook using LightCells mode
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the processed workbook
            workbook.Save(outputPath);
        }
    }
}
