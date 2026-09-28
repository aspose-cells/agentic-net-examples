// Title: Delete an external data connection from an Excel workbook by its description using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that searches the workbook's DataConnections for a ConnectionDescription matching a supplied value and removes that connection. | Demonstrate how to iterate over Workbook.DataConnections, identify a connection by its description, delete it, and save the modified .xlsx file.
// Common Searches: asp.net remove external data connection from Excel file using Aspose.Cells | c# find and delete workbook connection by description Aspose.Cells | how to programmatically delete a data connection in an .xlsx with Aspose.Cells | remove specific external connection from Excel workbook based on ConnectionDescription in C#
// Tags: Aspose.Cells external connection removal | C# delete workbook data connection | ExternalConnectionCollection filter by description | Excel .xlsx connection deletion Aspose | Aspose.Cells DataConnections manipulation

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads an Excel workbook, scans its DataConnections for a connection whose ConnectionDescription matches a given string, removes that connection if found, and saves the updated file.
class RemoveExternalConnectionByDescription
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Description of the connection to be removed
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
        if (indexToRemove != -1)
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
