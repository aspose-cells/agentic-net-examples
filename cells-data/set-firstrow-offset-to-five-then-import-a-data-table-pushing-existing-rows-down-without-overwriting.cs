// Title: Insert five blank rows at the top of an Excel sheet and import a DataTable starting at row 6 using Aspose.Cells for .NET
// Common Searches: how to add multiple empty rows at the beginning of an Excel worksheet with Aspose.Cells C# | import DataTable into Excel starting from row 6 without overwriting existing rows using Aspose.Cells | push existing rows down before writing a DataTable to an .xlsx file in C# | write column headers and data from a DataTable after inserting top rows with Aspose.Cells | save workbook after inserting rows and importing a DataTable in Aspose.Cells .NET
// Tags: insert rows at top Aspose.Cells C# | DataTable import with row offset Aspose.Cells | preserve existing worksheet data Aspose.Cells | write DataTable headers manually Aspose.Cells | save workbook to .xlsx Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, inserts five blank rows at the top to offset the first row, writes the DataTable's column names to the sixth row and its data rows starting at the seventh row, and saves the result as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (or add a new one)
                Worksheet sheet = workbook.Worksheets[0];

                // Insert five empty rows at the top to create a first‑row offset of 5
                sheet.Cells.InsertRows(0, 5);

                // Prepare a sample DataTable to import
                DataTable dt = new DataTable();
                dt.Columns.Add("ID", typeof(int));
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Score", typeof(double));

                // Add some rows to the DataTable
                dt.Rows.Add(1, "Alice", 85.5);
                dt.Rows.Add(2, "Bob", 92.0);
                dt.Rows.Add(3, "Charlie", 78.3);

                // Import the DataTable manually starting at row index 5 (the sixth row, after the offset)
                // Write column headers
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    sheet.Cells[5, col].PutValue(dt.Columns[col].ColumnName);
                }

                // Write data rows
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        sheet.Cells[6 + row, col].PutValue(dt.Rows[row][col]);
                    }
                }

                // Define output file path
                string outputPath = "Result.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
