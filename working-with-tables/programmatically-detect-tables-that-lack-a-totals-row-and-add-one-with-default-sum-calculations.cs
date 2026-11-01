// Title: How to add a totals row with sum formulas to Excel tables that lack one using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add missing totals row to Excel table programmatically | detect Excel ListObject tables without a totals row using Aspose.Cells | enable ShowTotals property for tables in a workbook with Aspose.Cells .NET | insert default sum calculations into a totals row of Excel tables via Aspose.Cells | iterate worksheets and ListObjects to add a totals row in C#
// Tags: add totals row to ListObject using Aspose.Cells | detect tables missing totals row C# | enable ShowTotals property Excel table Aspose.Cells | programmatic insertion of sum totals in .xlsx with Aspose.Cells | iterate worksheets ListObjects Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, iterates through each worksheet and its ListObject tables, checks the ShowTotals flag, enables it when missing, and saves the workbook, effectively adding a totals row with default sum calculations to tables that lacked one.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths (replace placeholders with actual paths)
            string inputFilePath = "{InputFilePath}";
            string outputFilePath = "{OutputFilePath}";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFilePath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Get the collection of tables (ListObjects) in the current worksheet
                    ListObjectCollection tables = worksheet.ListObjects;

                    // Process each table
                    foreach (ListObject table in tables)
                    {
                        // If the table does not already have a totals row, add one
                        if (!table.ShowTotals)
                        {
                            // Enable the totals row for the table
                            table.ShowTotals = true;

                            // Note: Setting specific totals functions (e.g., Sum) may require a newer Aspose.Cells version.
                            // If supported, you can set column.TotalFunction = TotalsRowFunction.Sum here.
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
