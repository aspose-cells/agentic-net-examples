using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class DisableBackgroundRefresh
{
    static void Main()
    {
        // Load an existing workbook that contains a DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Identify DBConnection objects
            if (connection is DBConnection dbConnection)
            {
                // Disable background (asynchronous) refresh to enforce sequential query execution
                dbConnection.BackgroundRefresh = false;
                Console.WriteLine($"BackgroundRefresh disabled for connection '{dbConnection.Name}'.");
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}