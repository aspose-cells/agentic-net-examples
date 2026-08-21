// Title: Get PivotTable External Connection Strings with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, walks through every worksheet and its PivotTables, calls GetSourceDataConnections() to fetch any ExternalConnection objects, and prints each connection's ConnectionString to the console. No workbook changes are made, but the file can be saved afterward.
// Keywords: Aspose.Cells | C# | PivotTable | ExternalConnection | GetSourceDataConnections | connection string | Excel audit | data source lookup
// Common Searches: Aspose.Cells get pivot table connection string | How to list external data connections of a PivotTable in C# | Retrieve pivot source connections using Aspose.Cells | C# read external connection from Excel pivot | Audit pivot table data sources Aspose
// Developer Intent: Extract and display the external data source connection strings used by each PivotTable in a workbook.
// Use Cases: Audit all PivotTables to confirm they reference approved data sources before publishing. | Generate a compliance report that lists every external connection string in an Excel file. | Detect and flag PivotTables that point to deprecated or insecure external databases.
// AI Prompts: Write C# code with Aspose.Cells that collects all external connection strings from PivotTables and writes them to a CSV file. | Show how to update the ConnectionString of a specific PivotTable's external data source using Aspose.Cells. | Provide a method to filter and list only PivotTables that have more than one external data connection.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsPivotConnectionAudit
{
    // Loads an Excel workbook, walks through every worksheet and its PivotTables, calls GetSourceDataConnections() to fetch any ExternalConnection objects, and prints each connection's ConnectionString to the console. No workbook changes are made, but the file can be saved afterward.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table with an external data connection
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables in the worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Retrieve the external data connections associated with the pivot table
                    ExternalConnection[] connections = pivot.GetSourceDataConnections();

                    // If there are any connections, display their connection strings
                    if (connections.Length > 0)
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}, PivotTable: {pivot.Name}");
                        for (int i = 0; i < connections.Length; i++)
                        {
                            Console.WriteLine($"  Connection {i + 1} String: {connections[i].ConnectionString}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}, PivotTable: {pivot.Name} has no external data connections.");
                    }
                }
            }

            // Optionally, save the workbook (no modifications made in this example)
            workbook.Save("output.xlsx");
        }
    }
}
