// Title: C# – Remove Empty Rows from an Aspose.Cells Worksheet Using a Bottom‑Up Loop
// Description: Demonstrates how to delete every completely blank row in an Aspose.Cells workbook. The example creates a workbook, adds sample data with intentional gaps, uses Cells.MaxDataRow and Cells.MaxDataColumn to locate the data range, iterates from the last row upward, checks each cell’s type for null, removes empty rows with Cells.DeleteRow, and saves the cleaned file as an XLSX document.
// Keywords: Aspose.Cells delete empty rows C# | remove blank rows Aspose.Cells .NET | Cells.MaxDataRow example | bottom up row deletion Aspose | DeleteRow method C# | clean worksheet Aspose.Cells | skip null cells Aspose.Cells
// Common Searches: how to delete empty rows in Aspose.Cells for .NET | remove blank rows from worksheet using C# Aspose.Cells | Aspose.Cells delete rows where all cells are null | iterate from bottom to top to delete rows Aspose.Cells | Aspose.Cells remove empty rows without shifting data
// Developer Intent: Eliminate all rows that contain no data from a worksheet while preserving the order of remaining rows.
// Use Cases: Sanitize CSV imports that contain sporadic empty lines before generating a report. | Trim placeholder rows added during dynamic data population in financial models. | Prepare a clean worksheet layout for export to Excel after programmatic row insertion.
// AI Prompts: Write C# code with Aspose.Cells that removes empty rows based on a single key column instead of the whole row. | Show an alternative method using Aspose.Cells Range objects or LINQ to filter out blank rows. | Explain how to adapt the loop to keep rows that have formulas but no visible values.

using System;
using Aspose.Cells;

// Demonstrates how to delete every completely blank row in an Aspose.Cells workbook. The example creates a workbook, adds sample data with intentional gaps, uses Cells.MaxDataRow and Cells.MaxDataColumn to locate the data range, iterates from the last row upward, checks each cell’s type for null, removes empty rows with Cells.DeleteRow, and saves the cleaned file as an XLSX document.
class RemoveEmptyRowsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the worksheet with some data and intentional blank rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            // Row 3 (index 2) is left blank
            cells["A4"].PutValue("Data2"); // Row 4 (index 3)
            // Row 5 (index 4) is left blank
            cells["A6"].PutValue("Data3"); // Row 6 (index 5)

            // Determine the last row that contains data (zero‑based index)
            int lastDataRow = cells.MaxDataRow;

            // Iterate from the bottom upwards so that row indices remain valid after deletions
            for (int row = lastDataRow; row >= 0; row--)
            {
                bool isEmpty = true;

                // Check each column up to the last column that contains data
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    // If any cell in the row is not null, the row is not empty
                    if (cells[row, col].Type != CellValueType.IsNull)
                    {
                        isEmpty = false;
                        break;
                    }
                }

                // Delete the row if it is empty
                if (isEmpty)
                {
                    cells.DeleteRow(row);
                }
            }

            // Save the workbook with the empty rows removed
            workbook.Save("RemovedEmptyRows.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
