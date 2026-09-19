// Title: Add a new row to an Excel ListObject table and populate it from a C# Dictionary using Aspose.Cells
// AI Prompts: Insert a new row after the last row of the first ListObject on a worksheet and write values from a Dictionary where each key matches a column header. | Resize the ListObject to include the newly added row and save the workbook to a different file path. | Skip any dictionary entries whose keys do not correspond to existing table headers while populating the row.
// Common Searches: how to insert a row into an Aspose.Cells ListObject from a C# dictionary | populate Excel table row with dictionary values using Aspose.Cells .NET | expand Aspose.Cells ListObject after adding a new row programmatically | match dictionary keys to Excel table column headers in C# Aspose.Cells
// Tags: insert row into ListObject Aspose.Cells | populate table from dictionary C# | resize ListObject after adding data Aspose.Cells | match dictionary keys to column headers Aspose.Cells | save modified workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads a workbook, retrieves the first ListObject, adds a new row after the table, fills cells by matching dictionary keys to column headers, expands the table range to include the new row, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found on the first worksheet.");
                return;
            }

            // Get the first table (ListObject) on the worksheet
            ListObject table = sheet.ListObjects[0];

            // Determine the index of the new row (first empty row after the table)
            int newWorksheetRowIndex = table.EndRow + 1;

            // Example dictionary containing column name → value pairs
            Dictionary<string, object> rowData = new Dictionary<string, object>()
            {
                { "Name", "John Doe" },
                { "Age", 28 },
                { "Country", "USA" }
            };

            // Populate the new row using the dictionary
            foreach (KeyValuePair<string, object> kvp in rowData)
            {
                // Find the column index that matches the dictionary key (header text)
                int columnIndex = -1;
                for (int col = table.StartColumn; col <= table.EndColumn; col++)
                {
                    string header = sheet.Cells[table.StartRow, col].StringValue;
                    if (header == kvp.Key)
                    {
                        columnIndex = col;
                        break;
                    }
                }

                // If the column was found, write the value into the new row cell
                if (columnIndex != -1)
                {
                    sheet.Cells[newWorksheetRowIndex, columnIndex].PutValue(kvp.Value);
                }
            }

            // Expand the table to include the newly added row
            table.Resize(table.StartRow, table.StartColumn, newWorksheetRowIndex, table.EndColumn, true);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
