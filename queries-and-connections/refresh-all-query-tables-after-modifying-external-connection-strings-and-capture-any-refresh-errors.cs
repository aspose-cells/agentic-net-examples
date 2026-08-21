// Title: Refresh all query tables after changing external connection strings and capture errors with Aspose.Cells for .NET
// Description: Load an Excel workbook, replace the connection string of every external connection used by query tables, call Workbook.RefreshAll, collect any refresh exceptions, and save the updated file. Includes robust error handling and console reporting.
// Keywords: Aspose.Cells refresh query tables | update external connection string .NET | Workbook.RefreshAll error handling | capture refresh errors C# | modify query table connection Aspose | Excel external connections programmatic | Aspose.Cells data source migration | C# Excel query table refresh
// Common Searches: change connection string for all query tables Aspose.Cells | refresh all query tables and get error messages .NET | Aspose.Cells Workbook.RefreshAll exception capture | update external connections in Excel using C# | how to log refresh failures with Aspose.Cells
// Developer Intent: Programmatically update every external connection string used by query tables, refresh the workbook, and log any refresh failures.
// Use Cases: Switch a workbook’s data source to a new Access or SQL database, refresh linked query tables, and record any failures for audit. | Automate nightly data pulls where the connection string changes daily and refresh errors must be captured for monitoring. | Migrate Excel reports to a new server, update all external connections in bulk, refresh the data, and ensure integrity by logging errors.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets, updates each query table's external connection string, calls RefreshAll, and stores any exceptions in a list. | Show how to handle Workbook.RefreshAll errors in Aspose.Cells for .NET and output them to the console or a log file. | Explain steps to verify that external connections were successfully updated before invoking RefreshAll using Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Load an Excel workbook, replace the connection string of every external connection used by query tables, call Workbook.RefreshAll, collect any refresh exceptions, and save the updated file. Includes robust error handling and console reporting.
class RefreshQueryTablesDemo
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "InputWorkbook.xlsx";
        string outputPath = "OutputWorkbook.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains query tables
            Workbook workbook = new Workbook(inputPath);

            // List to capture any refresh errors
            List<string> refreshErrors = new List<string>();

            // New connection string to be applied to all external connections
            string newConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\NewData\\Sample.accdb;";

            // Update connection strings of external connections used by query tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (QueryTable queryTable in sheet.QueryTables)
                {
                    ExternalConnection extConn = queryTable.ExternalConnection;
                    if (extConn != null)
                    {
                        extConn.ConnectionString = newConnectionString;
                    }
                }
            }

            // Refresh all external connections and query tables in the workbook
            try
            {
                workbook.RefreshAll();
            }
            catch (Exception ex)
            {
                refreshErrors.Add($"Error during workbook refresh: {ex.Message}");
            }

            // Optionally, display refresh errors
            if (refreshErrors.Count > 0)
            {
                Console.WriteLine("Refresh errors encountered:");
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
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
