using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RemoveExternalConnectionByDescription
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Description of the external connection to be removed
        string targetDescription = "Target connection description";

        // Get the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Find the index of the connection with the specified description
        int indexToRemove = -1;
        for (int i = 0; i < connections.Count; i++)
        {
            if (connections[i].ConnectionDescription == targetDescription)
            {
                indexToRemove = i;
                break;
            }
        }

        // Remove the connection if it was found
        if (indexToRemove >= 0)
        {
            connections.RemoveAt(indexToRemove);
            Console.WriteLine($"Removed external connection with description: \"{targetDescription}\"");
        }
        else
        {
            Console.WriteLine($"No external connection found with description: \"{targetDescription}\"");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}