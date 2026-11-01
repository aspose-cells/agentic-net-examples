// Title: Find cells containing the word “invoice” in workbook‑scoped named ranges with a case‑insensitive search using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to locate cells with specific text in workbook scoped named ranges | case insensitive string search across all named ranges in an Excel file using Aspose.Cells | retrieve addresses of cells that contain 'invoice' within named ranges with Aspose.Cells .NET | enumerate workbook scoped names and search for a keyword in Aspose.Cells C# | search for a word in Excel named ranges without caring about case using Aspose.Cells
// Tags: case-insensitive text search in workbook scoped named ranges Aspose.Cells | enumerate workbook scoped names C# Aspose.Cells | collect cell addresses from named ranges Aspose.Cells | filter string cells in Excel workbook Aspose.Cells | iterate named range cells Aspose.Cells C#

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, iterates through every workbook‑scoped named range, examines each string cell, performs a case‑insensitive comparison for the word "invoice", records matching cell addresses together with their worksheet names, and outputs the results.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to collect addresses of cells that contain the word "invoice"
            List<string> matches = new List<string>();

            // Iterate over all named ranges in the workbook
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Get the range that the name refers to
                Aspose.Cells.Range range = name.GetRange();

                // Scan each cell in the range
                foreach (Cell cell in range)
                {
                    // Check only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string value = cell.StringValue;
                        // Case‑insensitive search for "invoice"
                        if (!string.IsNullOrEmpty(value) &&
                            value.IndexOf("invoice", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Record the cell address with its sheet name
                            matches.Add($"{cell.Name} (Sheet: {cell.Worksheet.Name})");
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cells containing the word \"invoice\" (case‑insensitive) in workbook‑scoped named ranges:");
            foreach (string address in matches)
            {
                Console.WriteLine(address);
            }

            // If you need to save any changes, uncomment the line below
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
