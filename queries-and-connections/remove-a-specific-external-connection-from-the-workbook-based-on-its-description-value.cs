using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RemoveExternalConnectionByDescription
{
    static void Main()
    {
        // Path to the workbook that contains external connections
        string inputPath = "InputWorkbook.xlsx";
        // Path where the modified workbook will be saved
        string outputPath = "OutputWorkbook.xlsx";
        // Description value of the connection to be removed
        string targetDescription = "Obsolete connection";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Get the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate backwards so that removal does not affect the loop index
        for (int i = connections.Count - 1; i >= 0; i--)
        {
            ExternalConnection conn = connections[i];
            // Check the ConnectionDescription property
            if (conn.ConnectionDescription == targetDescription)
            {
                // Remove the connection at the found index
                connections.RemoveAt(i);
                Console.WriteLine($"Removed connection with description: \"{targetDescription}\" at index {i}");
            }
        }

        // Save the modified workbook
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to \"{outputPath}\"");
    }
}