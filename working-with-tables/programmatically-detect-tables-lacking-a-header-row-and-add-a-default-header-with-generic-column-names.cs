// Title: Add a default header row to Excel tables missing headers using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, iterates through each worksheet and ListObject, checks if ShowHeaderRow is false, sets it to true, and writes generic headers such as Column1, Column2 into the newly created header row. | Update the example so it accepts a custom header prefix (e.g., "Field") and an output file path via command‑line arguments, while still adding default headers to tables that lack a header row.
// Common Searches: how to programmatically add missing header rows to Excel tables with Aspose.Cells in C# | Aspose.Cells C# detect ListObject without header and insert default column names | C# iterate worksheets and ListObjects to enable ShowHeaderRow property | add generic column headers to Excel tables that lack a header using Aspose.Cells | set ShowHeaderRow = true and populate header cells in Aspose.Cells
// Tags: add default header to Excel ListObject Aspose.Cells | detect tables missing header row C# | enable ShowHeaderRow for ListObject programmatically | populate generic column names in Excel table | iterate worksheets and tables Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The program loads an input XLSX workbook, scans every worksheet and its ListObjects, enables the header row for any table where ShowHeaderRow is false, inserts generic column names (Column1, Column2, …) into the new header row, and saves the modified workbook to the specified output file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all tables (ListObjects) in the worksheet
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        // If the table does not have a header row, add one
                        if (!table.ShowHeaderRow)
                        {
                            table.ShowHeaderRow = true;

                            int headerRowIndex = table.StartRow;          // Row where the header will be placed
                            int firstColumnIndex = table.StartColumn;    // First column of the table
                            int columnCount = table.EndColumn - table.StartColumn + 1; // Number of columns in the table

                            // Insert generic column names (Column1, Column2, ...) into the header row
                            for (int col = 0; col < columnCount; col++)
                            {
                                string headerName = $"Column{col + 1}";
                                sheet.Cells[headerRowIndex, firstColumnIndex + col].PutValue(headerName);
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
