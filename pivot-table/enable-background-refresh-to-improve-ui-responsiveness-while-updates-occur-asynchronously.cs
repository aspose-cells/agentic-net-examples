using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class EnableBackgroundRefresh
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Enable background refresh for each external connection
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Refresh asynchronously in the background
            connection.BackgroundRefresh = true;
        }

        // Save the workbook with the modified connection settings
        workbook.Save("BackgroundRefreshEnabled.xlsx");
    }
}