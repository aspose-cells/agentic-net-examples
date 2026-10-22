// Title: Insert five rows at row 5 and import a DataTable into an Aspose.Cells worksheet without overwriting existing data
// AI Prompts: Add five blank rows starting at row index 5, then write a DataTable's column names and rows into the worksheet using Aspose.Cells for .NET. | Shift existing worksheet rows down by a specified offset and populate the newly created area with a DataTable, preserving all original content.
// Common Searches: Aspose.Cells insert rows before importing DataTable | C# how to add row offset when exporting DataTable to Excel with Aspose.Cells | push existing rows down in Aspose.Cells worksheet | import DataTable into specific row in Aspose.Cells without overwriting | Aspose.Cells insert rows and write DataTable headers
// Tags: insert rows offset Aspose.Cells | import DataTable to worksheet Aspose.Cells | shift existing rows down Aspose.Cells | write DataTable headers Aspose.Cells | save workbook after inserting rows Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, inserts five rows at index 5 to make space, writes the DataTable column names as headers, fills the rows with the DataTable values, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Example DataTable to import (replace with actual data source)
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("ID", typeof(int));
                dataTable.Columns.Add("Name", typeof(string));
                dataTable.Rows.Add(1, "Alice");
                dataTable.Rows.Add(2, "Bob");
                dataTable.Rows.Add(3, "Charlie");

                // Row offset where the import should start (zero‑based index)
                int startRow = 5;

                // Insert rows to make space for header + data
                int rowsToInsert = dataTable.Rows.Count + 1; // +1 for column headers
                sheet.Cells.InsertRows(startRow, rowsToInsert);

                // Write column headers
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    sheet.Cells[startRow, col].PutValue(dataTable.Columns[col].ColumnName);
                }

                // Write data rows
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        sheet.Cells[startRow + 1 + row, col].PutValue(dataTable.Rows[row][col]);
                    }
                }

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
