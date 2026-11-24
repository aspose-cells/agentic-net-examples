// Title: Delete a row from an Excel ListObject by primary key value using Aspose.Cells for .NET
// Common Searches: aspnet c# delete row from Excel table where Id = 12345 using Aspose.Cells | remove specific record from ListObject based on primary key with Aspose.Cells | how to locate and delete a table row by column value in an .xlsx file using Aspose.Cells | C# Aspose.Cells delete worksheet table row using primary key column | programmatically delete Excel table row by Id column in .NET
// Tags: Aspose.Cells delete ListObject row by column value | C# remove Excel table record using primary key | Aspose.Cells worksheet.Cells.DeleteRow with table | Excel .xlsx row deletion based on Id column | Aspose.Cells locate row in ListObject

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing workbook, checks for a ListObject, finds the column named 'Id', searches the table for the row whose Id matches a given value, deletes that row with worksheet.Cells.DeleteRow (which automatically updates the table range), and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one table
            if (worksheet.ListObjects.Count == 0)
                throw new Exception("No tables found in the worksheet.");

            // Access the first table (ListObject)
            var table = worksheet.ListObjects[0];

            // Primary‑key column name and value to identify the row to delete
            string primaryKeyColumnName = "Id";
            string primaryKeyValue = "12345";

            // Find the column offset of the primary‑key within the table
            int pkColumnOffset = -1;
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                if (table.ListColumns[i].Name.Equals(primaryKeyColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    pkColumnOffset = i;
                    break;
                }
            }
            if (pkColumnOffset == -1)
                throw new Exception($"Column '{primaryKeyColumnName}' not found in the table.");

            // Determine the start cell of the table's data range
            int dataStartRow = table.DataRange.FirstRow;
            int dataStartColumn = table.DataRange.FirstColumn;

            // Locate the row index (relative to the table) that matches the primary‑key value
            int targetRowIndex = -1;
            int totalRows = table.DataRange.RowCount; // number of data rows in the table
            for (int i = 0; i < totalRows; i++)
            {
                var cell = worksheet.Cells[dataStartRow + i, dataStartColumn + pkColumnOffset];
                if (cell.Value != null && cell.Value.ToString() == primaryKeyValue)
                {
                    targetRowIndex = i;
                    break;
                }
            }

            // Delete the row if found
            if (targetRowIndex >= 0)
            {
                int absoluteRow = dataStartRow + targetRowIndex;
                // Delete the row; the table range updates automatically
                worksheet.Cells.DeleteRow(absoluteRow);
            }
            else
            {
                throw new Exception($"Row with primary key value '{primaryKeyValue}' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
