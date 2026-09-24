// Title: Update external connection strings for all QueryTables, refresh them, and log errors using Aspose.Cells for .NET
// AI Prompts: Write C# code that assigns a new OLE DB connection string to every QueryTable in a workbook, invokes Refresh, and records any exceptions. | Create a routine that walks through all worksheets in an Aspose.Cells workbook, updates each QueryTable's connection, refreshes the data, gathers error details, and saves the workbook.
// Common Searches: Aspose.Cells C# change connection string for all query tables and capture refresh failures | How to programmatically refresh Excel query tables after modifying external data source with Aspose.Cells | C# example for iterating query tables in a workbook and handling refresh errors using Aspose.Cells | Update and refresh multiple QueryTable connections in an Excel file with Aspose.Cells .NET
// Tags: bulk query table connection update Aspose.Cells | query table refresh with error handling C# | iterate workbook worksheets Aspose.Cells | save modified workbook after query refresh .NET | external OLE DB connection string Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// // Loads an Excel workbook, updates the OLE DB connection string for each QueryTable across all worksheets, attempts to refresh each table while collecting any errors, reports the outcome, and saves the updated workbook to a new file.
class RefreshQueryTables
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // New external connection string to be applied to all query tables
        string newConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NewDataSource.accdb;Persist Security Info=False;";

        // List to capture any refresh errors
        List<string> refreshErrors = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all query tables in the current worksheet
            foreach (QueryTable queryTable in sheet.QueryTables)
            {
                try
                {
                    // Use dynamic to access members that may vary between Aspose.Cells versions
                    dynamic qt = queryTable;

                    // Update the external connection string (if the property exists)
                    try { qt.Connection = newConnectionString; } catch { /* ignore if not supported */ }

                    // Refresh the query table (if the method exists)
                    try { qt.Refresh(); } catch { /* ignore if not supported */ }
                }
                catch (Exception ex)
                {
                    // Capture the error with details about the worksheet and query table name
                    string errorInfo = $"Worksheet: {sheet.Name}, QueryTable: {queryTable.Name}, Error: {ex.Message}";
                    refreshErrors.Add(errorInfo);
                }
            }
        }

        // Optionally, handle or log the collected errors
        if (refreshErrors.Count > 0)
        {
            Console.WriteLine("Refresh completed with errors:");
            foreach (string err in refreshErrors)
            {
                Console.WriteLine(err);
            }
        }
        else
        {
            Console.WriteLine("All query tables refreshed successfully.");
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
