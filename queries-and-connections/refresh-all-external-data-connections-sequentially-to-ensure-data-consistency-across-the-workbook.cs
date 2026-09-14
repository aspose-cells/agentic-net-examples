// Title: Refresh every external data connection sequentially in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that opens an .xlsx file, loops through its ExternalConnection collection, sets RefreshOnLoad to true, turns off BackgroundRefresh, forces a synchronous refresh, recalculates all formulas, and saves the workbook. | Modify the given Aspose.Cells example so that each external connection is refreshed one after another before the workbook is written to disk. | Create a reusable C# method using Aspose.Cells that ensures all linked data sources are updated synchronously and then triggers a full formula recalculation.
// Common Searches: Aspose.Cells C# how to refresh external data connections before saving workbook | set RefreshOnLoad and turn off BackgroundRefresh for Excel connections using .NET | sequentially update multiple external connections in an .xlsx with Aspose.Cells | recalculate formulas after external connection refresh in Aspose.Cells | force synchronous data connection refresh in Aspose.Cells .NET
// Tags: Aspose.Cells external connection refresh | turn off background refresh Aspose.Cells | enable RefreshOnLoad property .NET | recalculate workbook formulas Aspose.Cells | save workbook after connection update

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads an Excel file, iterates over its ExternalConnection collection, enables RefreshOnLoad, disables background refresh for synchronous updates, recalculates all formulas, and saves the refreshed workbook.
class RefreshExternalConnections
{
    static void Main()
    {
        // Load the workbook that contains external data connections
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate through each connection and trigger a refresh
        foreach (ExternalConnection conn in connections)
        {
            // Ensure the connection is set to refresh when the workbook is opened
            conn.RefreshOnLoad = true;

            // Force synchronous refresh (no background processing)
            conn.BackgroundRefresh = false;
        }

        // Recalculate all formulas so that any data retrieved by the connections is applied
        workbook.CalculateFormula();

        // Save the workbook after all connections have been refreshed
        workbook.Save("output.xlsx");
    }
}
