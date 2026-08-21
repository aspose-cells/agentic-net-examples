// Title: List Power Query Formula Names in an Excel Workbook with Aspose.Cells for .NET
// Description: Loads a workbook, accesses its DataMashup.PowerQueryFormulas collection, prints each formula's Name to the console for verification, handles missing files and empty collections, and saves the workbook to a new file.
// Keywords: Aspose.Cells PowerQueryFormulas | list Power Query formulas C# | DataMashup enumeration .NET | read Power Query names from Excel | console output Power Query formula names | Aspose.Cells workbook save after read
// Common Searches: how to get Power Query formula names using Aspose.Cells | C# enumerate DataMashup PowerQueryFormulas | Aspose.Cells list Power Query queries in workbook | check for Power Query formulas before saving Excel file | exception handling Aspose.Cells file not found
// Developer Intent: Extract and display every Power Query formula name from an Excel file, then persist the workbook.
// Use Cases: Validate that imported workbooks contain the expected Power Query queries by listing their names. | Debug Power Query connections by outputting formula identifiers before performing refresh operations. | Maintain workbook lifecycle compliance by saving the file after read‑only operations.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, iterates over DataMashup.PowerQueryFormulas, and writes each formula.Name to the console. | Show how to detect an empty PowerQueryFormulas collection and provide a friendly console message before saving the workbook. | Explain the proper try‑catch pattern for FileNotFoundException and generic errors when loading a workbook and listing Power Query formulas.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsExamples
{
    // Loads a workbook, accesses its DataMashup.PowerQueryFormulas collection, prints each formula's Name to the console for verification, handles missing files and empty collections, and saves the workbook to a new file.
    public class ListPowerQueryFormulaNames
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the source file exists before attempting to load.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that may contain Power Query formulas.
                Workbook workbook = new Workbook(inputPath);

                // Access the DataMashup object which holds the PowerQueryFormulas collection.
                DataMashup mashup = workbook.DataMashup;

                // Verify that the collection exists and contains items.
                if (mashup?.PowerQueryFormulas != null && mashup.PowerQueryFormulas.Count > 0)
                {
                    Console.WriteLine("Power Query Formula Names:");
                    // Iterate through each PowerQueryFormula and output its Name.
                    foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                    {
                        Console.WriteLine(formula.Name);
                    }
                }
                else
                {
                    Console.WriteLine("No Power Query formulas found in the workbook.");
                }

                // Save the workbook (required by lifecycle rules, even if unchanged).
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
