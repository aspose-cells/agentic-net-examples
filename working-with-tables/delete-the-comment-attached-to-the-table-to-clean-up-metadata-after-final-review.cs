// Title: Delete a comment from an Excel table (ListObject) using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove comment from ListObject table | how to clear table comment in an .xlsx file with Aspose.Cells | programmatically delete Excel table comment using Aspose.Cells .NET | Aspose.Cells delete ListObject comment before saving workbook | remove metadata comment from Excel table in C# Aspose.Cells
// Tags: Aspose.Cells ListObject comment removal | C# clear Excel table comment | Aspose.Cells delete table comment .xlsx | programmatic comment cleanup Aspose.Cells | remove ListObject metadata comment C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates over each ListObject (table) on the first worksheet, clears the Comment property when it contains text, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Iterate through all tables (ListObjects) on the worksheet
                foreach (var table in sheet.ListObjects)
                {
                    // If the table has a comment, clear it to delete the comment
                    if (!string.IsNullOrEmpty(table.Comment))
                    {
                        table.Comment = string.Empty;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
