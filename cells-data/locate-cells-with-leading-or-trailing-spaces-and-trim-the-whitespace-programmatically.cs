// Title: Trim Leading and Trailing Spaces from Excel Cells using Aspose.Cells for .NET (C#)
// Description: Load a workbook, walk the used range of each worksheet, trim leading/trailing spaces from string cells, update only changed values, and save the cleaned file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | trim whitespace | remove spaces from Excel cells | string cell cleanup | Excel data normalization | used range iteration | Aspose.Cells for .NET | Excel workbook trimming
// Common Searches: Aspose.Cells remove leading spaces C# | trim trailing whitespace in Excel using Aspose | how to clean string values in a workbook with Aspose.Cells | C# code to trim spaces from Excel cells | Aspose.Cells whitespace cleanup example
// Developer Intent: Programmatically eliminate surrounding spaces from string cells in an Excel workbook with Aspose.Cells for .NET.
// Use Cases: Sanitize imported CSV data that contains accidental spaces before analysis. | Standardize user‑entered text in spreadsheets to ensure reliable lookups and reporting. | Prepare data for integration with external systems where exact string matching is required.
// AI Prompts: Write C# code using Aspose.Cells to iterate over all worksheets and trim leading/trailing spaces from every string cell, then save the workbook. | Show how to modify the loop to skip empty rows and process only non‑blank string cells while trimming whitespace. | Explain how to apply the same whitespace‑trimming logic to a specific range or column in a worksheet with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTrimWhitespace
{
    // Load a workbook, walk the used range of each worksheet, trim leading/trailing spaces from string cells, update only changed values, and save the cleaned file with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or iterate through all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Iterate over the used range of cells
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        string trimmed = original.Trim();

                        // Update the cell only if trimming changed the value
                        if (!original.Equals(trimmed))
                        {
                            cell.PutValue(trimmed);
                        }
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
