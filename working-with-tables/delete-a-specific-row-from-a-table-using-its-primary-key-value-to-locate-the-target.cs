// Title: Delete a specific Excel table row by primary key value using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that finds a row in a ListObject where a given primary‑key column matches a value, removes that row with Aspose.Cells, and then adjusts the table range. | Show how to programmatically locate an ID column in an Excel table, delete the matching row, and call the Resize method to keep the table headers intact using Aspose.Cells for .NET.
// Common Searches: c# remove specific row from Excel table by ID with Aspose.Cells | aspnet delete row from Excel ListObject using ID column Aspose.Cells | how to resize an Aspose.Cells table after deleting a row | find and delete row in Excel sheet based on column value using Aspose.Cells C# | Aspose.Cells delete table row without breaking table structure
// Tags: delete ListObject row Aspose.Cells | resize Excel table after row deletion Aspose.Cells | locate row by ID column C# | Aspose.Cells update table range | C# remove Excel table row by ID

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an XLSX workbook, accesses a ListObject named "Table1", identifies the row where the "ID" column equals a specified value, deletes that row, resizes the table to reflect the new range, and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Access the table (ListObject) by its name (replace "Table1" with your table's name)
            ListObject table = sheet.ListObjects["Table1"];
            if (table == null)
            {
                Console.WriteLine("Table \"Table1\" not found on the worksheet.");
                return;
            }

            // Define the primary key column name and the value to locate
            string primaryKeyColumnName = "ID";   // change to your PK column name
            int primaryKeyValue = 123;            // change to the PK value you want to delete

            // Find the column index of the primary key within the table header row
            int pkColumnIndex = -1;
            for (int col = table.StartColumn; col <= table.EndColumn; col++)
            {
                if (sheet.Cells[table.StartRow, col].StringValue.Equals(primaryKeyColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    pkColumnIndex = col;
                    break;
                }
            }

            if (pkColumnIndex == -1)
            {
                Console.WriteLine("Primary key column not found.");
                return;
            }

            // Locate the row that contains the specified primary key value
            int rowToDelete = -1;
            // Start from StartRow + 1 to skip the header row
            for (int row = table.StartRow + 1; row <= table.EndRow; row++)
            {
                Cell cell = sheet.Cells[row, pkColumnIndex];
                // Assuming integer primary key; adjust if needed
                if (cell.IntValue == primaryKeyValue)
                {
                    rowToDelete = row;
                    break;
                }
            }

            if (rowToDelete != -1)
            {
                // Delete the identified row from the worksheet
                sheet.Cells.DeleteRows(rowToDelete, 1);

                // Adjust the table range after row deletion
                int newEndRow = table.EndRow - 1;
                if (newEndRow >= table.StartRow)
                {
                    int totalRows = newEndRow - table.StartRow + 1;
                    int totalColumns = table.EndColumn - table.StartColumn + 1;
                    // The table has headers, so pass true for hasHeaders
                    table.Resize(table.StartRow, table.StartColumn, totalRows, totalColumns, true);
                }

                Console.WriteLine($"Row with {primaryKeyColumnName} = {primaryKeyValue} deleted.");
            }
            else
            {
                Console.WriteLine("Row with the specified primary key value not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
