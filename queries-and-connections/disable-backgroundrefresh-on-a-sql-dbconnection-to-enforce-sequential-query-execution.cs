// Title: Disable BackgroundRefresh on DBConnection objects in Aspose.Cells (C#) for sequential query execution
// Description: Loads an Excel workbook, loops through its external connections, finds each DBConnection (SQL, ODBC, OLE DB), sets BackgroundRefresh to false to enforce synchronous query processing, logs the change, and saves the updated file.
// Keywords: Aspose.Cells | DBConnection | BackgroundRefresh | disable background refresh | sequential query execution | C# | Excel external data connections | .NET | SQL connection | ODBC | OLE DB
// Common Searches: Aspose.Cells disable BackgroundRefresh C# | Set DBConnection.BackgroundRefresh false | Force sequential data refresh in Excel with Aspose.Cells | Turn off async refresh for SQL connections in Aspose.Cells | Iterate workbook.DataConnections and modify refresh settings
// Developer Intent: Turn off background refresh for all DBConnection objects so queries execute one after another.
// Use Cases: Ensure ordered data retrieval by disabling background refresh before saving the workbook. | Log each connection name and its BackgroundRefresh state for debugging purposes. | Apply the setting selectively by checking the connection's Name property.
// AI Prompts: Generate C# code using Aspose.Cells that disables BackgroundRefresh for every DBConnection in a workbook and saves the result. | Explain how the BackgroundRefresh property affects query execution order in Aspose.Cells and why setting it to false guarantees sequential processing. | Show robust error‑handling patterns when iterating over workbook.DataConnections and updating DBConnection properties.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, loops through its external connections, finds each DBConnection (SQL, ODBC, OLE DB), sets BackgroundRefresh to false to enforce synchronous query processing, logs the change, and saves the updated file.
    public class DisableBackgroundRefreshDemo
    {
        public static void Main()
        {
            // Load an existing workbook that contains a DBConnection.
            // Replace "input.xlsx" with the path to your workbook.
            Workbook workbook = new Workbook("input.xlsx");

            try
            {
                // Iterate through all external connections in the workbook.
                foreach (ExternalConnection conn in workbook.DataConnections)
                {
                    // Check if the connection is a DBConnection (SQL/ODBC/OLE DB).
                    if (conn is DBConnection dbConn)
                    {
                        // Disable background refresh to force synchronous (sequential) execution.
                        dbConn.BackgroundRefresh = false;

                        // Optionally, display the new setting for verification.
                        Console.WriteLine($"Connection \"{dbConn.Name}\" BackgroundRefresh set to {dbConn.BackgroundRefresh}");
                    }
                }

                // Save the modified workbook.
                // Replace "output.xlsx" with the desired output path.
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved with BackgroundRefresh disabled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
