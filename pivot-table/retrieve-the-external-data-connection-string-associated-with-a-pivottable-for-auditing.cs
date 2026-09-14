// Title: How to retrieve external data connection strings from a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that enumerates all external connection strings of a specified PivotTable in an Excel file. | Create a method that returns a list of source data connection strings for the first PivotTable on the first worksheet using Aspose.Cells. | Write a script that audits and prints the external connections of a PivotTable, then saves the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# get pivot table source connection string from workbook | How to list external data connections of a PivotTable using .NET | Retrieve connection strings for Excel pivot tables programmatically with Aspose | Audit pivot table external connections in C# Aspose.Cells example | Get source data connections of first pivot table in Excel file using Aspose.Cells
// Tags: Aspose.Cells pivot table external connections | C# retrieve pivot table connection string | Aspose.Cells GetSourceDataConnections method | audit Excel pivot table data sources .NET | list pivot table source connections Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// The example loads an existing workbook, accesses the first worksheet, checks for PivotTables, obtains the first PivotTable, calls GetSourceDataConnections() to fetch any external data connections, prints each connection string, and optionally saves the workbook after auditing.
class RetrievePivotTableConnectionString
{
    static void Main()
    {
        // Load an existing workbook that contains a pivot table with an external data connection
        Workbook workbook = new Workbook("PivotWorkbook.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Verify that the worksheet contains at least one pivot table
        if (sheet.PivotTables.Count > 0)
        {
            // Get the first pivot table in the worksheet
            PivotTable pivot = sheet.PivotTables[0];

            // Retrieve all external data connections associated with the pivot table
            ExternalConnection[] connections = pivot.GetSourceDataConnections();

            // Output the connection string(s) for auditing purposes
            if (connections.Length > 0)
            {
                for (int i = 0; i < connections.Length; i++)
                {
                    Console.WriteLine($"Connection {i + 1} String: {connections[i].ConnectionString}");
                }
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

        // Save the workbook (optional, if any changes were made)
        workbook.Save("PivotWorkbook_Audited.xlsx");
    }
}
