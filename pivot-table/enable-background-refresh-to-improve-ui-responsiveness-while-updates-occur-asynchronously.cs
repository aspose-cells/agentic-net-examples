using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using System;

class EnableBackgroundRefreshDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Check if the workbook contains any external connections
        if (workbook.DataConnections.Count > 0)
        {
            // Access the first external connection
            ExternalConnection connection = workbook.DataConnections[0];

            // Enable background (asynchronous) refresh
            connection.BackgroundRefresh = true;

            Console.WriteLine("BackgroundRefresh set to: " + connection.BackgroundRefresh);
        }
        else
        {
            Console.WriteLine("No external connections found in the workbook.");
        }

        // Save the workbook with the updated connection settings
        workbook.Save("BackgroundRefreshEnabled.xlsx");
    }
}