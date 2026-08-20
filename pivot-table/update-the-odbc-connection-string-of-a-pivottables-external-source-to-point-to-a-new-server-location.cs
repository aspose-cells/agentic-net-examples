// Title: Change PivotTable ODBC Connection String with Aspose.Cells for .NET
// Description: Loads a workbook, loops through its worksheets and PivotTables, finds external DBConnection objects, replaces their ODBC ConnectionString with a new server address, and saves the updated file.
// Keywords: Aspose.Cells | C# | PivotTable | ODBC connection string | external data source | DBConnection | Excel workbook | update server name | programmatic data connection | modify pivot source
// Common Searches: Aspose.Cells change PivotTable ODBC connection | C# update external data source for Excel PivotTable | set new server in PivotTable DBConnection string | programmatically modify PivotTable connection string .NET | batch update ODBC connections in Excel workbooks
// Developer Intent: Replace the ODBC connection string of a PivotTable’s external source so it points to a different database server.
// Use Cases: Repoint all PivotTables after migrating the database to a new host. | Automate workbook preparation for deployment by updating data source credentials. | Process a collection of reports to ensure they reference the correct server before distribution.
// AI Prompts: Generate C# code using Aspose.Cells that iterates through every PivotTable in a workbook and updates its ODBC ConnectionString to a specified server. | Show how to detect DBConnection objects within PivotTable source connections and safely assign a new connection string. | Explain best practices for error handling when saving a workbook after modifying external connections with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Loads a workbook, loops through its worksheets and PivotTables, finds external DBConnection objects, replaces their ODBC ConnectionString with a new server address, and saves the updated file.
    public class UpdatePivotTableOdbcConnection
    {
        // Entry point for the application
        public static void Main()
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
            // Path to the workbook that contains the pivot table with an ODBC connection
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // New ODBC connection string pointing to the new server location
            string newConnectionString = "Driver={SQL Server};Server=NewServerName;Database=MyDatabase;Trusted_Connection=Yes;";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables in the worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Get external data connections used by the pivot table
                    ExternalConnection[] connections = pivot.GetSourceDataConnections();

                    // Update each DBConnection's ConnectionString
                    foreach (ExternalConnection conn in connections)
                    {
                        if (conn is DBConnection dbConn)
                        {
                            dbConn.ConnectionString = newConnectionString;
                        }
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
