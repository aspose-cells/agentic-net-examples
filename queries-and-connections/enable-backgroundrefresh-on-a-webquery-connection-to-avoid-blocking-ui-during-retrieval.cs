using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class EnableBackgroundRefreshDemo
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Attempt to locate an existing WebQueryConnection.
        // If none exists, the demo will simply inform the user.
        if (workbook.DataConnections.Count > 0 && workbook.DataConnections[0] is WebQueryConnection webQuery)
        {
            // Enable background refresh to allow asynchronous data retrieval.
            webQuery.BackgroundRefresh = true;

            Console.WriteLine("BackgroundRefresh has been set to: " + webQuery.BackgroundRefresh);
        }
        else
        {
            Console.WriteLine("No WebQueryConnection found in the workbook. " +
                              "Add a web query connection first before setting BackgroundRefresh.");
        }

        // Save the workbook (the connection settings are persisted).
        workbook.Save("BackgroundRefreshEnabled.xlsx");
    }
}