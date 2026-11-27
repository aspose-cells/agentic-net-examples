// Title: Update Power Query source URL to a cloud storage location in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change Power Query source to Azure Blob storage URL | programmatically set Power Query data source URL in Excel using Aspose.Cells | replace Power Query source path with remote CSV link via Aspose.Cells .NET | modify DataMashup PowerQueryFormulaItem source value in C# example | set cloud storage URL for Power Query in an Excel file with Aspose.Cells API
// Tags: update Power Query source URL Aspose.Cells | modify DataMashup PowerQueryFormulaItem C# | set cloud storage URL for Excel Power Query | replace Power Query source path programmatically .NET | Aspose.Cells change Power Query data source URL

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace UpdatePowerQuerySource
{
    // The sample loads an Excel workbook, accesses its DataMashup to locate Power Query formulas, iterates through each PowerQueryFormulaItem, replaces any "Source" or "Path" values (or existing HTTP URLs) with a specified cloud storage URL, and saves the workbook as output.xlsx.
    class Program
    {
        /// <param name="args">
        /// args[0] – path to the source workbook (optional, defaults to "input.xlsx").
        /// args[1] – cloud storage URL to set as the new source (optional, defaults to a sample URL).
        /// </param>
        static void Main(string[] args)
        {
            // Determine input file and new URL
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string newSourceUrl = args.Length > 1
                ? args[1]
                : "https://mycloudstorage.blob.core.windows.net/data/sourcefile.csv";

            // Validate input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{inputPath}\" successfully.");

                // Access the DataMashup which holds Power Query formulas
                DataMashup mashup = workbook.DataMashup;
                if (mashup == null || mashup.PowerQueryFormulas == null || mashup.PowerQueryFormulas.Count == 0)
                {
                    Console.WriteLine("No Power Query formulas found in the workbook. No changes applied.");
                }
                else
                {
                    int modifiedCount = 0;

                    // Iterate through each Power Query formula
                    foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                    {
                        // Iterate through each item of the formula
                        foreach (PowerQueryFormulaItem item in formula.PowerQueryFormulaItems)
                        {
                            // The item that typically holds the source URL is named "Source"
                            // Some files may use "Path" or contain an HTTP/HTTPS string directly.
                            bool isPotentialSource =
                                string.Equals(item.Name, "Source", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(item.Name, "Path", StringComparison.OrdinalIgnoreCase) ||
                                (item.Value != null && item.Value.StartsWith("http", StringComparison.OrdinalIgnoreCase));

                            if (isPotentialSource)
                            {
                                Console.WriteLine($"Original value for item \"{item.Name}\": {item.Value}");
                                item.Value = newSourceUrl;
                                Console.WriteLine($"Updated value for item \"{item.Name}\": {item.Value}");
                                modifiedCount++;
                            }
                        }
                    }

                    Console.WriteLine($"{modifiedCount} Power Query source item(s) updated to the new URL.");
                }

                // Save the modified workbook
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory(),
                    "output.xlsx");

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with updated Power Query source to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
