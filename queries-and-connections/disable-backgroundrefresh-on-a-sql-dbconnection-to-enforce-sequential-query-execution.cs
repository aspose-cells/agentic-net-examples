// Title: Disable BackgroundRefresh on SQL DBConnection objects in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Iterate through Workbook.DataConnections and set DBConnection.BackgroundRefresh = false to make queries run one at a time in C# with Aspose.Cells. | Write C# code that loads an .xlsx file, locates all external DBConnection objects, turns off their background refresh, and saves the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# disable background refresh for SQL data connections | How to enforce sequential query execution for external DB connections in an Excel file using Aspose.Cells | Set BackgroundRefresh property to false for all DBConnection objects in a workbook with Aspose.Cells .NET | Turn off background refresh for external data connections in Aspose.Cells example
// Tags: background refresh off Aspose.Cells DBConnection | DBConnection.BackgroundRefresh false C# | ordered query execution Excel external connection Aspose.Cells | loop through workbook.DataConnections C# | Aspose.Cells external connection configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example loads an existing workbook (or creates a new one), loops through all external connections, disables BackgroundRefresh for each DBConnection to ensure queries run one at a time, and saves the modified workbook.
    public class DisableBackgroundRefreshDemo
    {
        // Entry point for the application
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Iterate through all external connections and disable background refresh for DB connections
            foreach (ExternalConnection connection in workbook.DataConnections)
            {
                if (connection is DBConnection dbConnection)
                {
                    dbConnection.BackgroundRefresh = false;
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
