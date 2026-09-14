// Title: How to recursively refresh nested pivot tables in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# method that accepts an Aspose.Cells PivotTable object and recursively refreshes its data cache and all dependent child pivot tables. | Demonstrate how to loop through every worksheet in a Workbook, call the recursive refresh for each pivot table hierarchy, and then save the updated workbook. | Provide error‑handled code that checks for the source file, loads the workbook, refreshes nested pivots, and writes the result to a new Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# refresh nested pivot tables in all worksheets | recursive refresh of dependent pivot tables using Aspose.Cells .NET | how to update pivot cache for child pivots programmatically with Aspose.Cells | C# code to refresh pivot table hierarchy in an Excel file
// Tags: Aspose.Cells recursive pivot refresh | refresh dependent pivot tables .NET | nested pivot tables Excel Aspose.Cells | pivot cache update C# Aspose.Cells | calculate pivot data programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, recursively refreshes each pivot table and its dependent child pivots by calling RefreshData and CalculateData, iterates through all worksheets to apply the refresh, and saves the updated workbook.
    public class RefreshNestedPivotTablesDemo
    {
        // Recursively refresh a pivot table and all its dependent (child) pivots
        private static void RefreshPivotAndChildren(PivotTable pivotTable)
        {
            try
            {
                // Refresh the pivot cache from the source data
                pivotTable.RefreshData();

                // Recalculate the pivot table values in the worksheet
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

        // Refresh all pivot tables in a worksheet, including nested ones
        private static void RefreshAllPivotTablesInWorksheet(Worksheet worksheet)
        {
            foreach (PivotTable pivotTable in worksheet.PivotTables)
            {
                RefreshPivotAndChildren(pivotTable);
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWithNestedPivots.xlsx";
            const string outputPath = "OutputWithRefreshedNestedPivots.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains the source data and pivot tables
                Workbook workbook = new Workbook(inputPath);

                // Refresh nested pivot tables for every worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    RefreshAllPivotTablesInWorksheet(sheet);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point required for compilation
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
