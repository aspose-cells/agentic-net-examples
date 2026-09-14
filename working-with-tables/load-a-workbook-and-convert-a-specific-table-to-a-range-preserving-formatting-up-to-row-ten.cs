// Title: Convert a named Excel table to a regular range and retain only the first ten rows with Aspose.Cells for .NET
// AI Prompts: Load the workbook 'input.xlsx' using Aspose.Cells, locate the ListObject named 'MyTable', invoke ConvertToRange to turn it into a plain range while keeping its visual style, then delete every row after the tenth data row and save the result as 'output.xlsx'. | In C#, use Aspose.Cells to transform a specific worksheet table into a cell range, preserve the original formatting, truncate the range to the first ten rows, and write the modified workbook.
// Common Searches: Aspose.Cells C# convert ListObject MyTable to range and keep first 10 rows | How to delete rows after converting an Excel table to range with Aspose.Cells .NET | Preserve cell formatting when converting a named table to range using Aspose.Cells | Limit rows of a converted Excel table to ten rows in C# Aspose.Cells example
// Tags: convert ListObject to range Aspose.Cells | retain formatting during table conversion .NET | remove rows beyond row ten Aspose.Cells | load workbook modify table rows C# | Aspose.Cells keep first ten rows table

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The code loads 'input.xlsx', finds the ListObject called 'MyTable', converts it to a regular range while preserving its formatting, deletes any rows beyond the 10th data row, and saves the updated workbook as 'output.xlsx'.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the table (ListObject) by name
            ListObject table = worksheet.ListObjects["MyTable"];
            if (table == null)
            {
                Console.WriteLine("Table 'MyTable' not found.");
                return;
            }

            // Store table boundaries before conversion
            int startRow = table.StartRow;
            int totalRows = table.DataRange.RowCount; // correct way to get row count

            // Convert the table to a regular range while preserving formatting
            table.ConvertToRange();

            // Delete rows beyond the 10th row (preserve rows 0‑9)
            int rowsToKeep = 10;
            if (totalRows > rowsToKeep)
            {
                int rowsToDelete = totalRows - rowsToKeep;
                worksheet.Cells.DeleteRows(startRow + rowsToKeep, rowsToDelete, true);
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
