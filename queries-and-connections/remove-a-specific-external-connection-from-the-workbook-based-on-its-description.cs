using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RemoveExternalConnectionByDescription
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "InputWorkbook.xlsx";

        // Description value of the connection to be removed
        string targetDescription = "Obsolete data source";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Get the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate backwards to safely remove items by index
        for (int i = connections.Count - 1; i >= 0; i--)
        {
            ExternalConnection conn = connections[i];

            // Check if the connection's description matches the target
            if (conn.ConnectionDescription == targetDescription)
            {
                // Remove the matching connection
                connections.RemoveAt(i);
            }
        }

        // Save the modified workbook
        workbook.Save("OutputWorkbook.xlsx");
    }
}