// Title: Delete a Table Row by Primary Key with Aspose.Cells for .NET (C#)
// Description: A C# sample that opens an Excel workbook, accesses the first ListObject on the initial sheet, determines the column containing the primary key, searches the data rows for a matching key value, deletes the located row using Cells.DeleteRow (shifting cells upward), and saves the modified file. Includes validation for missing files, absent tables, and undefined key columns.
// Keywords: Aspose.Cells delete row | C# Excel table row removal | ListObject delete by ID | primary key row deletion | Aspose.Cells .NET example | Excel worksheet DeleteRow method | remove record from Excel table | delete row by column value C# | Aspose.Cells table manipulation
// Common Searches: Aspose.Cells C# delete row where ID equals value | How to remove a record from an Excel ListObject using a key column | C# code to find and delete a table row in Excel with Aspose.Cells | Delete specific row in Excel table based on cell content
// Developer Intent: Identify and eliminate the row in an Excel table whose primary‑key column matches a given value.
// Use Cases: Purging a customer entry (ID = 5) from a sales report before distribution. | Removing an outdated product SKU from an inventory table during data cleanup. | Eliminating duplicate rows that share the same unique identifier in a generated analytics workbook.
// AI Prompts: Write C# code with Aspose.Cells that deletes a ListObject row where the 'ID' column equals a supplied integer and saves the workbook. | Enhance the sample with comprehensive error handling for scenarios such as missing workbook, no tables on the sheet, or non‑existent primary‑key column. | Show how to delete multiple rows that match a collection of primary‑key values using Aspose.Cells in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableRowDeletion
{
    // A C# sample that opens an Excel workbook, accesses the first ListObject on the initial sheet, determines the column containing the primary key, searches the data rows for a matching key value, deletes the located row using Cells.DeleteRow (shifting cells upward), and saves the modified file. Includes validation for missing files, absent tables, and undefined key columns.
    class Program
    {
        static void Main()
        {
            // Paths for input and output workbooks
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Primary key value to locate the row that should be deleted
            var primaryKeyValue = 5; // example: delete row where ID = 5

            // Name of the column that holds the primary key (case‑insensitive)
            string primaryKeyColumnName = "ID";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0]; // assuming the table is on the first sheet

                // Get the first table (ListObject) on the worksheet
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables found on the worksheet.");
                    return;
                }
                ListObject table = worksheet.ListObjects[0];

                // Determine the zero‑based index of the primary key column within the table
                int pkColumnIndex = -1;
                for (int col = 0; col < table.ListColumns.Count; col++)
                {
                    if (string.Equals(table.ListColumns[col].Name, primaryKeyColumnName, StringComparison.OrdinalIgnoreCase))
                    {
                        pkColumnIndex = col;
                        break;
                    }
                }

                if (pkColumnIndex == -1)
                {
                    Console.WriteLine($"Primary key column \"{primaryKeyColumnName}\" not found in the table.");
                    return;
                }

                // Calculate the first and last data row indices (skip header if present)
                int dataStartRow = table.StartRow + (table.ShowHeaderRow ? 1 : 0);
                int dataEndRow = table.EndRow; // inclusive

                // Locate the row that matches the primary key value
                int rowToDelete = -1;
                for (int row = dataStartRow; row <= dataEndRow; row++)
                {
                    Cell pkCell = worksheet.Cells[row, table.StartColumn + pkColumnIndex];

                    if (pkCell.Value != null && pkCell.Value.Equals(primaryKeyValue))
                    {
                        rowToDelete = row;
                        break;
                    }

                    // Fallback: compare as string (handles mismatched types)
                    if (pkCell.Value != null && pkCell.Value.ToString().Equals(primaryKeyValue.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        rowToDelete = row;
                        break;
                    }
                }

                if (rowToDelete == -1)
                {
                    Console.WriteLine($"No row with primary key value \"{primaryKeyValue}\" was found.");
                }
                else
                {
                    // Delete the identified row and shift cells up
                    worksheet.Cells.DeleteRow(rowToDelete, true);
                    Console.WriteLine($"Row {rowToDelete + 1} (primary key = {primaryKeyValue}) deleted.");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
