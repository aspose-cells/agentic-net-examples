// Title: C# – Delete a Pivot Table by Name Using Aspose.Cells for .NET
// Description: A C# utility that loads an Excel workbook with Aspose.Cells, locates a pivot table by its Name, removes it via PivotTableCollection.RemoveAt, and saves the result. Includes file‑existence validation, automatic output folder creation, and console feedback.
// Keywords: Aspose.Cells delete pivot table | remove pivot table by name C# | PivotTableCollection.RemoveAt example | programmatic Excel pivot table deletion | Aspose.Cells .NET utility
// Common Searches: delete pivot table by name Aspose.Cells | remove specific pivot table C# | Aspose.Cells PivotTableCollection RemoveAt usage | how to programmatically delete Excel pivot tables | C# console app delete pivot table
// Developer Intent: The developer needs to identify a pivot table by its name in an Excel file and delete it programmatically, then write the updated workbook to a new location.
// Use Cases: Strip temporary pivot tables from auto‑generated reports before distribution. | Clean up workbooks in a data‑processing pipeline where pivot tables become obsolete. | Provide a command‑line tool that accepts input file, pivot table name, and output file to remove unwanted pivot tables.
// AI Prompts: Generate a C# method that deletes a named pivot table using Aspose.Cells, with file checks and Save. | Write robust error‑handling code for removing a pivot table by name from all worksheets in a workbook. | Create a console application that takes input path, pivot table name, and output path to delete the specified pivot table.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableDemo
{
    // A C# utility that loads an Excel workbook with Aspose.Cells, locates a pivot table by its Name, removes it via PivotTableCollection.RemoveAt, and saves the result. Includes file‑existence validation, automatic output folder creation, and console feedback.
    public static class PivotTableUtility
    {
        /// <param name="inputFile">Path to the source Excel file.</param>
        /// <param name="pivotTableName">Name of the pivot table to delete.</param>
        /// <param name="outputFile">Path where the modified workbook will be saved.</param>
        public static void DeletePivotTableByName(string inputFile, string pivotTableName, string outputFile)
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook from the given file
                Workbook workbook = new Workbook(inputFile);

                bool removed = false;

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    PivotTableCollection pivots = sheet.PivotTables;

                    // Search for the pivot table by name
                    for (int i = 0; i < pivots.Count; i++)
                    {
                        if (pivots[i].Name.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Remove the pivot table using its index
                            pivots.RemoveAt(i);
                            removed = true;
                            break; // Exit the inner loop once removed
                        }
                    }

                    if (removed)
                        break; // Exit the outer loop if removal succeeded
                }

                if (!removed)
                {
                    Console.WriteLine($"Pivot table \"{pivotTableName}\" not found.");
                }

                // Ensure output directory exists
                string outDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                // Save the modified workbook to the specified output path
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Default parameters
            string inputFile = "Input.xlsx";
            string pivotTableName = "PivotTable1";
            string outputFile = "Output.xlsx";

            // Override with command‑line arguments if provided
            if (args.Length >= 3)
            {
                inputFile = args[0];
                pivotTableName = args[1];
                outputFile = args[2];
            }

            PivotTableUtility.DeletePivotTableByName(inputFile, pivotTableName, outputFile);
        }
    }
}
