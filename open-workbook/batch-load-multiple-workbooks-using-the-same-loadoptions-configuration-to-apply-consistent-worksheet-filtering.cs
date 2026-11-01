// Title: Batch load multiple Excel workbooks and filter worksheets by name prefix using Aspose.Cells for .NET
// Common Searches: aspnet load several .xlsx files and keep only worksheets that start with 'Data' using Aspose.Cells | c# batch process Excel workbooks to remove non‑matching worksheet names with Aspose.Cells | how to apply the same LoadOptions to multiple workbooks in Aspose.Cells .NET | filter worksheets by prefix while iterating over a list of Excel files in C# | save filtered Excel workbooks with a suffix using Aspose.Cells
// Tags: batch loading of Excel workbooks with Aspose.Cells | prefix‑based worksheet removal in C# | consistent LoadOptions usage across multiple workbooks | saving filtered workbooks with custom filename suffix | automated worksheet filtering Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example iterates over a collection of Excel file paths, loads each workbook with Aspose.Cells, removes any worksheet whose name does not start with "Data", and saves the filtered workbook to a new file that includes a "_filtered" suffix.
class Program
{
    static void Main()
    {
        try
        {
            // List of workbook file paths to process
            List<string> workbookPaths = new List<string>
            {
                "Book1.xlsx",
                "Book2.xlsx",
                "Book3.xlsx"
            };

            foreach (string path in workbookPaths)
            {
                try
                {
                    // Verify that the source file exists
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"Source file not found: {path}");
                        continue;
                    }

                    // Load the workbook (no load filter)
                    Workbook wb = new Workbook(path);

                    // Remove worksheets whose name does NOT start with "Data"
                    for (int i = wb.Worksheets.Count - 1; i >= 0; i--)
                    {
                        Worksheet ws = wb.Worksheets[i];
                        if (!ws.Name.StartsWith("Data", StringComparison.OrdinalIgnoreCase))
                        {
                            wb.Worksheets.RemoveAt(i);
                        }
                    }

                    // Save the filtered workbook to a new file
                    string outputPath = Path.GetFileNameWithoutExtension(path) + "_filtered.xlsx";
                    wb.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Filtered workbook saved to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{path}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
