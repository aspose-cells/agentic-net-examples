// Title: Refresh an ODBC-linked query table in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file, verifies the presence of query tables, refreshes the ODBC query table, and saves the updated workbook with Aspose.Cells. | Create a method that validates the input workbook path, calls Workbook.RefreshAll() to update all query tables, handles any refresh exceptions, and ensures the output directory exists before saving. | Develop a console application that loads a workbook, checks for at least one query table, refreshes external data sources programmatically, and logs errors while writing the refreshed file.
// Common Searches: Aspose.Cells .NET how to refresh ODBC query table in existing Excel file | C# program to refresh all query tables in a workbook and save changes | Refresh Excel query tables using Aspose.Cells without opening Excel UI | Handle missing input workbook when refreshing query tables with Aspose.Cells
// Tags: Workbook.RefreshAll ODBC query table | Aspose.Cells load workbook refresh external data | C# refresh Excel query tables programmatically | save refreshed workbook Aspose.Cells | error handling query table refresh .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQueryTableRefresh
{
    // The example loads an existing Excel workbook, checks for query tables on the first worksheet, refreshes all linked ODBC query tables using Workbook.RefreshAll(), creates the output directory if necessary, and saves the refreshed workbook to a new file while handling missing files and refresh errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "InputWorkbook.xlsx";
                string outputPath = "OutputWorkbook.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Assume the query table is on the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one query table
                if (sheet.QueryTables.Count == 0)
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                    return;
                }

                // Refresh all query tables in the workbook (covers the target table)
                try
                {
                    workbook.RefreshAll();
                }
                catch (Exception refreshEx)
                {
                    Console.WriteLine($"Error during query table refresh: {refreshEx.Message}");
                    return;
                }

                // Ensure the output directory exists (if a directory is specified)
                string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with refreshed data
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
