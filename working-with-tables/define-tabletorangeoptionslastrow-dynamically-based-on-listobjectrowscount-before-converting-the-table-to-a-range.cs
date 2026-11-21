// Title: Dynamically set TableToRangeOptions.LastRow from ListObject row count when converting an Excel table to a range with Aspose.Cells for .NET
// Common Searches: how to set TableToRangeOptions.LastRow using ListObject.Rows.Count in Aspose.Cells C# | convert Excel table to range with dynamic last row Aspose.Cells .NET example | Aspose.Cells C# TableToRangeOptions LastRow property based on table data range | C# code to change Excel ListObject to range with calculated last row using Aspose.Cells | Aspose.Cells .NET convert ListObject to range without hard‑coded row numbers
// Tags: TableToRangeOptions LastRow dynamic assignment | Aspose.Cells ListObject to range conversion | C# Excel table range conversion using Aspose.Cells | dynamic last row calculation for TableToRangeOptions | Aspose.Cells convert table to range .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example creates or loads an Excel workbook, retrieves the first ListObject, sets TableToRangeOptions.LastRow to the table's data range row count minus one, converts the table into a regular range using those options, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple workbook with a table if it does not.
                if (!File.Exists(inputPath))
                {
                    var tempWorkbook = new Workbook();
                    var tempWorksheet = tempWorkbook.Worksheets[0];

                    // Populate sample data.
                    tempWorksheet.Cells["A1"].PutValue("Header1");
                    tempWorksheet.Cells["B1"].PutValue("Header2");
                    tempWorksheet.Cells["A2"].PutValue("Data1");
                    tempWorksheet.Cells["B2"].PutValue("Data2");

                    // Add a table covering the data range (including header).
                    tempWorksheet.ListObjects.Add(0, 0, 2, 2, true);
                    tempWorkbook.Save(inputPath);
                }

                // Load the workbook from the input file.
                var workbook = new Workbook(inputPath);

                // Access the first worksheet.
                var worksheet = workbook.Worksheets[0];

                // Retrieve the first table (ListObject) on the worksheet.
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables found in the worksheet.");
                    return;
                }

                var table = worksheet.ListObjects[0];

                // Configure options for converting the table to a range.
                var options = new TableToRangeOptions
                {
                    // Set the last row based on the table's data range row count.
                    LastRow = table.DataRange.RowCount - 1
                };

                // Convert the table to a range using the configured options.
                table.ConvertToRange(options);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
