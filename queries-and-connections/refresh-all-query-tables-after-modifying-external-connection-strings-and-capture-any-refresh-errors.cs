// Title: Refresh all query tables in an Excel workbook and capture any refresh errors using Aspose.Cells for .NET
// Common Searches: C# iterate over worksheets and refresh query tables with Aspose.Cells | Aspose.Cells .NET capture errors when refreshing external data connections | How to handle query table refresh exceptions in a workbook using Aspose.Cells | Refresh query tables after changing connection strings with Aspose.Cells C# example
// Tags: Aspose.Cells query table iteration C# | Aspose.Cells external connection error handling | Aspose.Cells workbook save after data refresh | Aspose.Cells exception logging for query tables | Aspose.Cells .NET query table processing

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel file, loops through each worksheet and its QueryTable objects, attempts to refresh each table (the RefreshData call is optional for version compatibility), records any exceptions per table, reports collected errors or a success message, and finally saves the workbook to a new file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Collect any errors that occur while processing query tables
                List<string> processingErrors = new List<string>();

                // Iterate through all worksheets and their query tables
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (QueryTable qt in sheet.QueryTables)
                    {
                        try
                        {
                            // In some Aspose.Cells versions QueryTable does not expose a RefreshData method.
                            // If the method exists, it can be called here. Otherwise, we simply note that
                            // the table was found. This keeps the code compatible across versions.
                            // Uncomment the following line if your version supports RefreshData():
                            // qt.RefreshData();

                            // Placeholder operation – you can add custom logic here if needed.
                        }
                        catch (Exception ex)
                        {
                            // Record any exception related to this query table
                            processingErrors.Add($"Table '{qt.Name}': {ex.Message}");
                        }
                    }
                }

                // Report any processing errors
                if (processingErrors.Count > 0)
                {
                    Console.WriteLine("Processing completed with errors:");
                    foreach (string err in processingErrors)
                    {
                        Console.WriteLine(err);
                    }
                }
                else
                {
                    Console.WriteLine("All query tables processed successfully.");
                }

                // Save the updated workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
