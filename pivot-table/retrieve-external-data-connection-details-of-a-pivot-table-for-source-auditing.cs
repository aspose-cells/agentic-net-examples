// Title: C# – Retrieve PivotTable External Data Connections with Aspose.Cells
// Description: Loads a workbook, accesses the first worksheet, checks for PivotTables, and uses PivotTable.GetSourceDataConnections to list each ExternalConnection (name, class type, source type, command, connection string, description). Ideal for auditing or troubleshooting pivot table data sources, then saves the workbook.
// Keywords: Aspose.Cells PivotTable external connections | GetSourceDataConnections C# | audit pivot table data source | list Excel external connections .NET | retrieve pivot table connection string | Aspose.Cells example external data | Excel pivot table source auditing
// Common Searches: How to list external connections of a PivotTable using Aspose.Cells | Aspose.Cells C# get source data connections from pivot table | Audit Excel pivot table data sources with .NET | Retrieve connection string of a pivot table in C# | Aspose.Cells GetSourceDataConnections example
// Developer Intent: Extract and display all external data connection details used by a PivotTable for auditing, compliance, or troubleshooting purposes.
// Use Cases: Create an audit report of every external data source referenced by PivotTables before workbook distribution. | Validate that PivotTables use approved connection strings and source types in a compliance workflow. | Log connection details to diagnose refresh failures or data‑source mismatches.
// AI Prompts: Generate a C# snippet that iterates through all PivotTables in a workbook and logs each external connection's name, type, and connection string using Aspose.Cells. | Explain how to update the connection string of a PivotTable's external data source with Aspose.Cells for .NET. | Show code that filters PivotTables without external connections and outputs a warning message.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// Loads a workbook, accesses the first worksheet, checks for PivotTables, and uses PivotTable.GetSourceDataConnections to list each ExternalConnection (name, class type, source type, command, connection string, description). Ideal for auditing or troubleshooting pivot table data sources, then saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook that contains a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Get reference to the first pivot table
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Retrieve external data connections associated with the pivot table
        ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

        // Output connection details for auditing
        if (connections.Length == 0)
        {
            Console.WriteLine("The pivot table does not use any external data connections.");
        }
        else
        {
            foreach (ExternalConnection conn in connections)
            {
                Console.WriteLine($"Connection Name       : {conn.Name}");
                Console.WriteLine($"Class Type            : {conn.ClassType}");
                Console.WriteLine($"Source Type           : {conn.SourceType}");
                Console.WriteLine($"Command               : {conn.Command}");
                Console.WriteLine($"Connection String     : {conn.ConnectionString}");
                Console.WriteLine($"Connection Description: {conn.ConnectionDescription}");
                Console.WriteLine();
            }
        }

        // Save the workbook (optional – can be used to persist any changes)
        workbook.Save("AuditResult.xlsx");
    }
}
