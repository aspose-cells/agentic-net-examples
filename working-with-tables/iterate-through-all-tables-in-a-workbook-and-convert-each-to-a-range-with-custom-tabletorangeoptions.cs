// Title: Convert all Excel tables (ListObjects) to ranges in every worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates through each worksheet, and calls ListObject.ConvertToRange on every table while preserving formatting. | Write a C# program that safely converts all ListObjects in a workbook to regular ranges, handling missing input files and creating the output folder if needed. | Create a script that processes an Excel workbook with Aspose.Cells, converts tables to ranges in reverse order to avoid collection changes, and saves the modified file.
// Common Searches: Aspose.Cells C# convert all tables in a workbook to ranges | How to remove ListObject tables and keep data in Excel using Aspose.Cells .NET | Batch convert Excel tables to ranges with Aspose.Cells and preserve formulas | Iterate worksheets and convert ListObjects to ranges in C# Aspose.Cells example | Aspose.Cells TableToRangeOptions default behavior for converting tables
// Tags: Aspose.Cells convert ListObject to range | C# batch table-to-range conversion Excel | Aspose.Cells retain cell formatting during table conversion | convert tables in reverse order Aspose.Cells | ensure output path exists Aspose.Cells workbook save

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an input.xlsx workbook, verifies its existence, then loops through each worksheet. For each sheet it iterates the ListObjects collection backwards and calls ListObject.ConvertToRange, which transforms the table into a normal range while keeping formatting and formulas. Errors for individual tables are caught and reported. The program also creates the output directory if it does not exist, saves the modified workbook as output.xlsx, and writes status messages to the console.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Convert each table (ListObject) in every worksheet to a range
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate backwards because converting removes the table from the collection
                for (int i = sheet.ListObjects.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        ListObject table = sheet.ListObjects[i];
                        // Convert using default options (preserves formatting, formulas, etc.)
                        table.ConvertToRange();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to convert table on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
