// Title: C# custom function to find a value across all worksheets and return SheetName!CellAddress with Aspose.Cells
// Common Searches: Aspose.Cells C# find cell containing specific text in any worksheet | how to get sheet name and cell address for a lookup value using Aspose.Cells | search entire workbook for a string and return reference in C# | C# iterate worksheets and locate value with Aspose.Cells FindOptions | retrieve cell address of matching value across multiple Excel sheets in .NET
// Tags: Aspose.Cells workbook-wide value lookup | C# find cell address across worksheets | FindOptions LookInType.Values usage | search Excel workbook for string with Aspose.Cells | return SheetName!CellReference from Find method | iterate worksheets for value match in .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, defines a target string, and calls a custom FindValueInWorkbook function. The function iterates through each worksheet, uses Aspose.Cells Find with FindOptions (searching values and using Contains matching) to locate the first matching cell, and returns the address formatted as "SheetName!CellName". The result is printed and the workbook is saved.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Value to search for
                string searchValue = "Target";

                // Perform the lookup
                string? foundAddress = FindValueInWorkbook(workbook, searchValue);

                // Output the result
                if (!string.IsNullOrEmpty(foundAddress))
                    Console.WriteLine($"Value found at: {foundAddress}");
                else
                    Console.WriteLine("Value not found in any worksheet.");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Searches all worksheets for the specified value and returns the address (SheetName!CellName)
        static string? FindValueInWorkbook(Workbook wb, string value)
        {
            // Configure find options to search cell values only
            FindOptions options = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains // change to Exact if needed
            };

            foreach (Worksheet ws in wb.Worksheets)
            {
                // Search within the current worksheet
                Cell cell = ws.Cells.Find(value, null, options);
                if (cell != null)
                {
                    return $"{ws.Name}!{cell.Name}";
                }
            }

            // Not found
            return null;
        }
    }
}
