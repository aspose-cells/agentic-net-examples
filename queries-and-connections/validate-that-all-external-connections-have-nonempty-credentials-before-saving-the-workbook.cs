// Title: Ensure every external data connection has credentials before saving an Aspose.Cells workbook in C#
// AI Prompts: Loop through workbook.DataConnections, verify each ExternalConnection.Credentials is not CredentialsMethodType.None, and raise InvalidOperationException with the connection name if a missing credential is found. | Add a pre‑save validation method to an Aspose.Cells C# project that checks all external connections for non‑empty credentials and aborts the Save call when any connection lacks authentication. | Create a utility function that accepts a Workbook, inspects its ExternalConnection objects, and returns a boolean indicating whether all connections have proper credentials, throwing an error otherwise.
// Common Searches: Aspose.Cells C# verify external connection credentials before workbook.save() | how to throw error when external data connection has no credentials in Aspose.Cells | C# code to validate DataConnections credentials in an Excel file using Aspose | prevent saving Aspose.Cells workbook if any external connection lacks authentication | check for empty credentials in Aspose.Cells external connections C# example
// Tags: external connection credential validation Aspose.Cells | check DataConnections credentials C# | throw InvalidOperationException missing external credentials | pre‑save external connection check Aspose.Cells | validate external data connections before workbook save

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// // This program creates (or loads) a Workbook, iterates through its DataConnections, throws an InvalidOperationException if any connection's Credentials are set to None, and saves the workbook as ValidatedWorkbook.xlsx.
class ValidateExternalConnections
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Validate that the connection has credentials set (not the default 'None')
            if (connection.Credentials == CredentialsMethodType.None)
            {
                // Throw an exception if any connection lacks credentials
                throw new InvalidOperationException(
                    $"External connection '{connection.Name}' has empty credentials.");
            }
        }

        // All connections are valid; save the workbook
        workbook.Save("ValidatedWorkbook.xlsx");
    }
}
