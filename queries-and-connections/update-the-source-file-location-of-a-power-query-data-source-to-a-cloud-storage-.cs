using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class UpdatePowerQuerySource
{
    static void Main()
    {
        // Load the workbook that contains the Power Query data source
        Workbook workbook = new Workbook("input.xlsx");

        // New cloud storage URL to set as the source location
        string newCloudUrl = "https://mycloudstorage.blob.core.windows.net/data/sourcefile.xlsx";

        // Iterate through all external data connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // If the connection uses a file-based source, update its SourceFile property
            if (!string.IsNullOrEmpty(connection.SourceFile))
            {
                connection.SourceFile = newCloudUrl;
                Console.WriteLine($"Updated SourceFile for connection '{connection.Name}' to {newCloudUrl}");
            }

            // For web query connections, also update the Url property
            if (connection is WebQueryConnection webConn)
            {
                webConn.Url = newCloudUrl;
                Console.WriteLine($"Updated Url for WebQueryConnection '{webConn.Name}' to {newCloudUrl}");
            }
        }

        // Save the workbook with the updated data source locations
        workbook.Save("output.xlsx");
    }
}