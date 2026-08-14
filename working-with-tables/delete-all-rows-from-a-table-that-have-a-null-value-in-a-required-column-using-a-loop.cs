// Title: Aspose.Cells for .NET (C#) – Remove rows where a required column is null or blank
// Description: The example builds a workbook, fills column A with sample entries, iterates from the last row upward, evaluates each cell in the mandatory column for null, DBNull or an empty string, deletes rows that meet the condition using Cells.DeleteRow, and writes the result to RowsDeleted.xlsx.
// Keywords: Aspose.Cells | .NET | C# | DeleteRow | null values | empty cells | required column | reverse loop | Excel worksheet | MaxDataRow | data cleanup
// Common Searches: Aspose.Cells delete rows with null values C# | remove blank rows from Excel worksheet using .NET | loop delete rows where column is empty Aspose.Cells | C# Aspose.Cells delete rows based on required column | how to purge rows with missing data in Excel via Aspose
// Developer Intent: Programmatically eliminate rows that lack a mandatory value in a specific column.
// Use Cases: Sanitize imported CSV/Excel data by discarding records missing a key field before further processing. | Generate clean reports where rows without an identifier must be omitted. | Automate validation of Excel sheets in ETL pipelines by removing incomplete rows on the fly.
// AI Prompts: Generate C# code with Aspose.Cells that deletes rows where column B contains null, DBNull, or an empty string, iterating from the bottom to keep indexes stable. | Show an Aspose.Cells .NET snippet that removes rows with missing mandatory values and saves the workbook as an XLSX file. | Create a reusable method that accepts a Worksheet and a column index, then deletes all rows with null or blank cells in that column using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example builds a workbook, fills column A with sample entries, iterates from the last row upward, evaluates each cell in the mandatory column for null, DBNull or an empty string, deletes rows that meet the condition using Cells.DeleteRow, and writes the result to RowsDeleted.xlsx.
    public class DeleteRowsWithNullInRequiredColumn
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created and rows with null/empty values deleted successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (column A is the required column)
            // Row 0 – Header
            cells["A1"].PutValue("Name");
            // Rows with data; some rows have null/empty values in the required column
            cells["A2"].PutValue("Alice");   // valid
            cells["A3"].PutValue("");        // null/empty – should be deleted
            cells["A4"].PutValue("Bob");     // valid
            cells["A5"].PutValue(null);      // null – should be deleted
            cells["A6"].PutValue("Charlie"); // valid

            // Determine the index of the required column (0‑based, column A)
            int requiredColumnIndex = 0;

            // Loop from the last data row upwards to avoid index shifting after deletions
            for (int row = cells.MaxDataRow; row >= 0; row--)
            {
                // Retrieve the cell value; it can be null, DBNull, or an empty string
                object cellValue = cells[row, requiredColumnIndex].Value;

                // Check for null or empty string (treated as null for this scenario)
                bool isNullOrEmpty = cellValue == null ||
                                     (cellValue is string s && string.IsNullOrEmpty(s)) ||
                                     (cellValue is DBNull);

                if (isNullOrEmpty)
                {
                    // Delete the entire row
                    cells.DeleteRow(row);
                }
            }

            // Save the workbook to a file
            workbook.Save("RowsDeleted.xlsx", SaveFormat.Xlsx);
        }
    }
}
