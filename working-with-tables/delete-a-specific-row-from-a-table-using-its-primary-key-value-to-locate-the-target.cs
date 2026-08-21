// Title: Delete a Table Row by Primary Key with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, accesses the first worksheet and its ListObject, scans the primary‑key column for a specified value, removes the matching data row using Cells.DeleteRow, and saves the updated file.
// Keywords: Aspose.Cells delete row | C# delete Excel table row | ListObject remove row by key | Aspose.Cells primary key | Excel table row deletion .NET | Cells.DeleteRow example | Aspose.Cells find row by value
// Common Searches: Aspose.Cells delete row by primary key C# | How to remove a specific row from an Excel ListObject using Aspose.Cells | C# code to delete table row based on ID column | Delete Excel table row programmatically with Aspose.Cells | Find and delete row in Excel table using primary key value
// Developer Intent: Locate the row that matches a given primary‑key value in an Excel table and delete it.
// Use Cases: Purging obsolete records from a master data table | Applying a physical delete after confirming record existence | Synchronizing two spreadsheets by eliminating rows absent in the source file | User‑initiated removal of a selected entry in a data‑entry form | Automated cleanup of temporary rows generated during import
// AI Prompts: Write C# code using Aspose.Cells to locate a row in a ListObject where a specified column equals a given key and delete that row. | Show an Aspose.Cells example that scans a table for a primary‑key value, removes the matching row, and updates the table range. | Explain how to adapt the deletion logic when the primary‑key column is not the first column in the table. | Provide error‑handling code for missing tables or absent key values when deleting rows with Aspose.Cells. | Demonstrate how to delete multiple rows that match a list of primary‑key values in an Excel worksheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDeleteRowByPrimaryKey
{
    // Loads an Excel workbook, accesses the first worksheet and its ListObject, scans the primary‑key column for a specified value, removes the matching data row using Cells.DeleteRow, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                    return;
                }

                // Retrieve the first table; you can also get by name: worksheet.ListObjects["TableName"]
                ListObject table = worksheet.ListObjects[0];

                // Define the primary key value you want to delete
                object primaryKeyToDelete = 5; // example primary key value

                // Determine the column index of the primary key (assumed first column of the table)
                int primaryKeyColumnIndex = table.StartColumn;

                // Determine the first data row index (skip header if present)
                int firstDataRow = table.StartRow + (table.ShowHeaderRow ? 1 : 0);
                int rowToDelete = -1;

                // Scan the table rows to locate the row with the matching primary key
                for (int row = firstDataRow; row <= table.EndRow; row++)
                {
                    object cellValue = worksheet.Cells[row, primaryKeyColumnIndex].Value;
                    if (cellValue != null && cellValue.Equals(primaryKeyToDelete))
                    {
                        rowToDelete = row;
                        break;
                    }
                }

                // If a matching row was found, delete it and update references
                if (rowToDelete != -1)
                {
                    worksheet.Cells.DeleteRow(rowToDelete, true);
                    Console.WriteLine($"Row {rowToDelete + 1} (primary key = {primaryKeyToDelete}) deleted.");
                }
                else
                {
                    Console.WriteLine($"No row found with primary key = {primaryKeyToDelete}.");
                }

                // Save the modified workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
