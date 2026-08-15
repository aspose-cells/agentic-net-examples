// Title: C# – Retrieve External ODBC Connection String from a Pivot Table using Aspose.Cells
// Description: Loads a workbook, accesses the first worksheet, checks for pivot tables, calls GetSourceDataConnections() on the first pivot, and writes the ODBC connection string to the console. The workbook can be saved unchanged.
// Keywords: Aspose.Cells C# pivot external connection | GetSourceDataConnections | ODBC connection string pivot table | read pivot source data Aspose | extract external connection Aspose.Cells
// Common Searches: how to get ODBC connection string from a pivot table in .NET | Aspose.Cells retrieve external data connections for pivot | C# code to read pivot table source connection | log external ODBC string using Aspose.Cells
// Developer Intent: Extract and display the ODBC connection string that a pivot table uses as its external data source.
// Use Cases: Verify that a pivot table points to the intended ODBC data source before processing. | Capture connection strings for audit trails or troubleshooting workbook links. | Compare external data sources across multiple pivot tables in a batch operation.
// AI Prompts: Show how to modify the retrieved ODBC connection string and reassign it to the pivot table with Aspose.Cells. | Add comprehensive error handling for missing or multiple external connections in a pivot table. | Generate code that iterates through all pivot tables in a workbook and prints each external ODBC connection string.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// Loads a workbook, accesses the first worksheet, checks for pivot tables, calls GetSourceDataConnections() on the first pivot, and writes the ODBC connection string to the console. The workbook can be saved unchanged.
class RetrievePivotExternalConnectionString
{
    static void Main()
    {
        // Load an existing workbook that contains a pivot table with an external ODBC connection
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one pivot table
        if (sheet.PivotTables.Count > 0)
        {
            // Get the first pivot table
            PivotTable pivot = sheet.PivotTables[0];

            // Retrieve external data connections associated with the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            if (connections.Length > 0)
            {
                // Log the connection string of the first external connection
                Console.WriteLine("External ODBC Connection String: " + connections[0].ConnectionString);
            }
            else
            {
                Console.WriteLine("No external data connections found for the pivot table.");
            }
        }
        else
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
        }

        // Save the workbook (optional if no changes were made)
        workbook.Save("output.xlsx");
    }
}
