// Title: Recursively Refresh Nested Pivot Tables in Excel Using Aspose.Cells for .NET (C#)
// Description: This C# sample shows how to load an Excel workbook with Aspose.Cells, walk through each worksheet and its pivot tables, and update the pivot cache and calculated values for a pivot table and any linked child tables. It leverages RefreshData, CalculateData, and GetDependentPivotTables, includes file‑existence checks and exception handling, and writes the changes back to a new file.
// Keywords: Aspose.Cells | C# | refresh pivot tables | nested pivot tables | recursive pivot refresh | GetDependentPivotTables | RefreshData | CalculateData | Excel workbook .NET | pivot cache update
// Common Searches: Aspose.Cells refresh nested pivot tables | C# recursive pivot table refresh example | How to update child pivot tables in Excel with Aspose | GetDependentPivotTables usage .NET | RefreshData and CalculateData Aspose.Cells
// Developer Intent: Update every pivot table and its dependent tables so they reflect the latest source data.
// Use Cases: Automate pivot table updates after source data changes before distributing a report. | Process workbooks that contain multiple levels of linked pivot tables in a reporting pipeline. | Ensure data consistency when source ranges are modified in large Excel files. | Integrate a pivot‑refresh routine into existing Aspose.Cells automation scripts. | Handle missing input files gracefully while performing bulk workbook transformations.
// AI Prompts: Write C# code using Aspose.Cells that walks through all worksheets and recursively refreshes each pivot table and its dependent tables, with proper error handling. | Show how to use GetDependentPivotTables to locate child pivot tables and refresh them after calling RefreshData and CalculateData. | Explain how to embed a recursive pivot‑refresh function into a larger Aspose.Cells workflow that processes multiple Excel files.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This C# sample shows how to load an Excel workbook with Aspose.Cells, walk through each worksheet and its pivot tables, and update the pivot cache and calculated values for a pivot table and any linked child tables. It leverages RefreshData, CalculateData, and GetDependentPivotTables, includes file‑existence checks and exception handling, and writes the changes back to a new file.
    public class RefreshNestedPivotTables
    {
        // Recursively refresh a pivot table and all its dependent (child) pivot tables
        private static void RefreshPivotAndChildren(PivotTable pivotTable)
        {
            try
            {
                // Refresh the pivot cache from the source data
                pivotTable.RefreshData();

                // Recalculate the pivot table values
                pivotTable.CalculateData();

                // Get dependent pivot tables that use this pivot as a data source
                PivotTable[] children = pivotTable.GetDependentPivotTables();

                // Recursively refresh each dependent pivot table
                foreach (PivotTable child in children)
                {
                    RefreshPivotAndChildren(child);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error refreshing pivot table '{pivotTable.Name}': {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook that contains the pivot tables
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and their pivot tables
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        RefreshPivotAndChildren(pivot);
                    }
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point required for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
