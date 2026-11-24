// Title: Remove a comment from a ListObject (Excel table) using Aspose.Cells for .NET (C#)
// Common Searches: how to clear a table comment with Aspose.Cells C# | Aspose.Cells remove ListObject comment programmatically | delete comment attached to Excel table using .NET SDK | set ListObject.Comment to null Aspose.Cells example | remove metadata comment from Excel table in C#
// Tags: Aspose.Cells ListObject comment removal | C# clear Excel table comment property | Aspose.Cells delete table comment .NET | Excel table metadata cleanup Aspose.Cells | set ListObject.Comment null C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, accesses the first worksheet, finds the ListObject named "Table1", checks for an attached comment, clears it by assigning null to the Comment property, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the table named "Table1"
                ListObject table = worksheet.ListObjects["Table1"];
                if (table == null)
                {
                    Console.WriteLine("Table 'Table1' not found.");
                }
                else
                {
                    // Delete the comment attached to the table, if it exists
                    if (!string.IsNullOrEmpty(table.Comment))
                    {
                        table.Comment = null;
                    }
                }

                // Save the modified workbook
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
