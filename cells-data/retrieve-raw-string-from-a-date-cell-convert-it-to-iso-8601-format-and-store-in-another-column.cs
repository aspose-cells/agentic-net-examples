// Title: Convert Excel Date Cells to ISO‑8601 Strings and Write to Adjacent Column with Aspose.Cells for .NET
// Description: Load an Excel workbook using Aspose.Cells, iterate through a source column, detect DateTime cells, format each date with the round‑trip "o" pattern (ISO‑8601), place the formatted string in a target column, copy non‑date values unchanged, and save the updated file.
// Keywords: Aspose.Cells C# date conversion | Excel ISO 8601 format | convert Excel date to string .NET | write values to another column Aspose | DateTimeValue to ISO string | Excel to JSON date format | C# Excel date handling | Aspose.Cells example | ISO 8601 round‑trip format | load and save workbook Aspose
// Common Searches: Aspose.Cells convert Excel date to ISO 8601 | C# read Excel date cell and output ISO string | write formatted date to another column using Aspose.Cells | detect DateTime cells in Aspose.Cells | save workbook after modifying cells Aspose | ISO 8601 round‑trip format C# Excel | Aspose.Cells example for date formatting
// Developer Intent: Read dates from a column, format them as ISO‑8601 strings, and write the results to a neighboring column.
// Use Cases: Export Excel data to APIs that require ISO‑8601 timestamps | Create audit sheets that show original dates alongside ISO representations | Prepare data for JSON or XML serialization where dates must follow ISO‑8601 | Standardize date formats across regional workbooks | Generate reports that need UTC‑compatible timestamps
// AI Prompts: Write Aspose.Cells C# code to read dates from column A, convert each to ISO‑8601 using the "o" format, and store in column B, preserving non‑date cells. | Provide a reusable method that accepts a workbook path, source and target column indexes, converts DateTime cells to ISO‑8601 strings, and saves the file. | Explain how to handle time‑zone offsets when formatting Excel dates to ISO‑8601 with Aspose.Cells. | Show how to batch‑process multiple worksheets for date conversion to ISO‑8601. | Demonstrate error handling for cells that are not dates in an Aspose.Cells date‑conversion routine.

using System;
using Aspose.Cells;

namespace AsposeCellsDateConversion
{
    // Load an Excel workbook using Aspose.Cells, iterate through a source column, detect DateTime cells, format each date with the round‑trip "o" pattern (ISO‑8601), place the formatted string in a target column, copy non‑date values unchanged, and save the updated file.
    class Program
    {
        static void Main()
        {
            // Create a workbook instance and load an existing Excel file
            Workbook workbook = new Workbook("input.xlsx"); // load rule

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define source (date) column index and target column index
            int sourceColumn = 0; // Column A (zero‑based)
            int targetColumn = 1; // Column B (zero‑based)

            // Iterate through all used rows in the source column
            int maxRow = cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                Cell dateCell = cells[row, sourceColumn];

                // Ensure the cell actually contains a DateTime value
                if (dateCell.Type == CellValueType.IsDateTime)
                {
                    // Retrieve the DateTime value
                    DateTime dt = dateCell.DateTimeValue;

                    // Convert to ISO 8601 string (round‑trip format)
                    string isoString = dt.ToString("o"); // e.g., 2023-05-15T00:00:00.0000000Z

                    // Store the ISO string in the target column
                    cells[row, targetColumn].PutValue(isoString);
                }
                else
                {
                    // If the cell is not a DateTime, you may choose to handle it differently.
                    // Here we simply copy the original string (if any) to the target column.
                    cells[row, targetColumn].PutValue(dateCell.StringValue);
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx"); // save rule
        }
    }
}
