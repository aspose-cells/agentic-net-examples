// Title: Delete a row from an Excel ListObject by primary key value using Aspose.Cells for .NET
// AI Prompts: Locate the 'ID' column in the first ListObject, find the data row whose value matches a given integer, delete that row with Aspose.Cells, and save the workbook. | Programmatically remove an Excel table row identified by a primary‑key column and persist the changes using C# and Aspose.Cells. | Search for a column named 'ID' in a worksheet table, delete the matching row, and write the updated file with Aspose.Cells.
// Common Searches: C# Aspose.Cells delete table row where column value equals specific integer | How to remove a ListObject row by primary key in Aspose.Cells .NET | Aspose.Cells delete Excel row based on ID column value | Remove specific data row from Excel table using Aspose.Cells example
// Tags: Aspose.Cells delete ListObject row by column value | C# remove Excel table row using Aspose.Cells | primary key row deletion in Excel with Aspose.Cells | save workbook after row removal Aspose.Cells | search column index in ListObject Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads an Excel workbook, identifies the first ListObject, finds the row where the 'ID' column matches a supplied primary‑key integer, deletes that worksheet row (the table updates automatically), and saves the modified workbook to a new file.
class DeleteTableRowByPrimaryKey
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Primary key value to locate the row that must be deleted
        int primaryKeyValue = 123; // example value

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count == 0)
                throw new Exception("No tables (ListObjects) found on the worksheet.");

            ListObject table = worksheet.ListObjects[0];

            // Name of the primary key column in the table
            string primaryKeyColumnName = "ID";

            // Find the index of the primary key column within the table
            int pkColumnIndex = -1;
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                if (table.ListColumns[i].Name.Equals(primaryKeyColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    pkColumnIndex = i;
                    break;
                }
            }

            if (pkColumnIndex == -1)
                throw new Exception($"Primary key column \"{primaryKeyColumnName}\" not found in the table.");

            // Locate the row index (within the table) that matches the primary key value
            int rowIndexInTable = -1;
            int dataRowCount = table.DataRange.RowCount; // number of data rows (excluding header)

            for (int i = 0; i < dataRowCount; i++)
            {
                // Data rows start after the header row
                int worksheetRow = table.StartRow + i + 1;
                int worksheetColumn = table.StartColumn + pkColumnIndex;

                Cell cell = worksheet.Cells[worksheetRow, worksheetColumn];

                // Compare the cell value with the target primary key value
                if (cell.IntValue == primaryKeyValue)
                {
                    rowIndexInTable = i;
                    break;
                }
            }

            // Delete the row if found
            if (rowIndexInTable >= 0)
            {
                // Delete the corresponding worksheet row; the table adjusts automatically
                int worksheetRowToDelete = table.StartRow + rowIndexInTable + 1;
                worksheet.Cells.DeleteRows(worksheetRowToDelete, 1);
                Console.WriteLine($"Row with primary key {primaryKeyValue} deleted.");
            }
            else
            {
                Console.WriteLine($"No row with primary key value {primaryKeyValue} was found.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
