// Title: Using Aspose.Cells ICellsDataTable to stream a database query into a worksheet starting at row zero (C#)
// AI Prompts: Write C# code that implements the ICellsDataTable interface to stream rows from a SqlDataReader directly into an Aspose.Cells worksheet beginning at cell A1, then save the workbook as an XLSX file. | Show how to map a DataTable's column names to the first row of an Aspose.Cells worksheet and fill subsequent rows with data using the Cells API in C#. | Generate a method that accepts an IDataReader, creates an ICellsDataTable implementation, and writes the data to the first worksheet of a new Aspose.Cells workbook starting at row zero.
// Common Searches: Aspose.Cells C# import DataTable starting at row 0 with column headers | Implement ICellsDataTable to write SQL query results to Excel using Aspose.Cells | C# example of streaming database rows into Aspose.Cells worksheet at top row | How to save a DataTable to XLSX with Aspose.Cells preserving header row
// Tags: ICellsDataTable implementation for Excel export | import DataTable to Aspose.Cells worksheet | write database query results to XLSX using Aspose.Cells | populate worksheet header row with DataTable columns | stream IDataReader into Aspose.Cells workbook | C# Aspose.Cells export to XLSX

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Aspose.Cells Workbook, builds a sample DataTable, writes the column names to the first row (row 0) of the first worksheet, fills subsequent rows with the DataTable values, ensures the output directory exists, and saves the workbook as Output.xlsx while handling potential errors.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                var workbook = new Workbook();

                // 2. Get the first worksheet
                var sheet = workbook.Worksheets[0];

                // 3. Prepare a DataTable (replace with real DB logic if needed)
                var dt = new DataTable();
                try
                {
                    // Sample columns
                    dt.Columns.Add("Id", typeof(int));
                    dt.Columns.Add("Name", typeof(string));
                    dt.Columns.Add("CreatedDate", typeof(DateTime));

                    // Sample rows
                    dt.Rows.Add(1, "Alice", DateTime.Now);
                    dt.Rows.Add(2, "Bob", DateTime.Now.AddDays(-1));
                    dt.Rows.Add(3, "Charlie", DateTime.Now.AddDays(-2));
                }
                catch (Exception dtEx)
                {
                    Console.Error.WriteLine($"Error creating DataTable: {dtEx.Message}");
                    return;
                }

                // 4. Import the DataTable into the worksheet (include column headers)
                try
                {
                    // Write column headers
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        sheet.Cells[0, col].PutValue(dt.Columns[col].ColumnName);
                    }

                    // Write data rows
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        for (int col = 0; col < dt.Columns.Count; col++)
                        {
                            sheet.Cells[row + 1, col].PutValue(dt.Rows[row][col]);
                        }
                    }
                }
                catch (Exception importEx)
                {
                    Console.Error.WriteLine($"Error importing data to worksheet: {importEx.Message}");
                    return;
                }

                // 5. Save the workbook
                const string outputPath = "Output.xlsx";

                // Ensure the output directory exists
                var outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.Error.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
