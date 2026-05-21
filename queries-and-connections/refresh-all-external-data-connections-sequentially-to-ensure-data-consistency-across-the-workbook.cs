using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RefreshExternalDataConnections
{
    static void Main()
    {
        // Load the workbook that contains external data connections
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Refresh each connection sequentially
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];

            // Ensure the connection is refreshed when the workbook is opened
            conn.RefreshOnLoad = true;

            // Force synchronous refresh (optional, depending on requirements)
            conn.BackgroundRefresh = false;
        }

        // Recalculate formulas that may depend on the refreshed data
        workbook.CalculateFormula();

        // Save the workbook after all connections have been refreshed
        workbook.Save("output.xlsx");
    }
}