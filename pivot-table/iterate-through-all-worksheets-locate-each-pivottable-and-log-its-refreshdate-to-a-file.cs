// Title: Log the RefreshDate of each PivotTable in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that opens an Excel file, iterates through all worksheets, reads every PivotTable's RefreshDate, and writes worksheet name, pivot name, and date to a log file. | Enhance the PivotRefreshDateLogger to also capture each PivotTable's cache ID and include it in the same log output. | Create a reusable method that returns a list of objects containing worksheet name, pivot table name, RefreshDate, and cache ID for further processing.
// Common Searches: aspocells get pivot table refreshdate c# example | how to write pivot table metadata to a log file using Aspose.Cells | enumerate all pivot tables in a workbook and retrieve their refresh timestamps .net | c# Aspose.Cells iterate worksheets and extract pivot properties | log pivot cache id and refresh date from Excel using Aspose.Cells
// Tags: Aspose.Cells read pivot refresh date | C# log pivot table metadata to file | enumerate worksheets pivot tables Aspose.Cells | export pivot refresh timestamps .NET | capture pivot cache identifier Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, loops through every worksheet, extracts each PivotTable's RefreshDate (and optionally cache ID), and writes the worksheet name, pivot name, and date to a log file.
    public class PivotRefreshDateLogger
    {
        // Adjust these paths as needed
        private const string InputFilePath = "input.xlsx";
        private const string LogFilePath = "PivotRefreshDates.log";

        public static void Run()
        {
            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(InputFilePath))
                {
                    Console.WriteLine($"Input file not found: {InputFilePath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(InputFilePath);

                // Open a StreamWriter to write the log file
                using (StreamWriter writer = new StreamWriter(LogFilePath, false))
                {
                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Access the collection of pivot tables in the current worksheet
                        PivotTableCollection pivotTables = sheet.PivotTables;

                        // If there are no pivot tables, continue to the next worksheet
                        if (pivotTables == null || pivotTables.Count == 0)
                            continue;

                        // Iterate through each pivot table
                        for (int i = 0; i < pivotTables.Count; i++)
                        {
                            PivotTable pivot = pivotTables[i];

                            // Retrieve the RefreshDate property
                            DateTime refreshDate = pivot.RefreshDate;

                            // Write worksheet name, pivot table name, and refresh date to the log
                            writer.WriteLine($"Worksheet: {sheet.Name}, PivotTable: {pivot.Name}, RefreshDate: {refreshDate}");
                        }
                    }
                }

                Console.WriteLine($"Pivot refresh dates logged to {LogFilePath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotRefreshDateLogger.Run();
        }
    }
}
