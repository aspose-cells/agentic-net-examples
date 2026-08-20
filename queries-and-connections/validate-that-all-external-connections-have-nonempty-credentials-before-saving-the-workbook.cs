// Title: Validate External Connection Credentials Before Saving an Aspose.Cells Workbook (.NET)
// Description: Shows how to loop through a Workbook’s DataConnections in C#, confirm each ExternalConnection has a non‑empty Credentials value, raise an InvalidOperationException for any missing credentials, and save the workbook only after validation succeeds.
// Keywords: Aspose.Cells | C# external connection validation | DataConnections credentials | Aspose.Cells .NET | validate workbook external data source | CredentialsMethodType.None | throw exception missing credentials | save workbook after validation
// Common Searches: Aspose.Cells check external connection credentials | C# validate DataConnections before saving workbook | ensure external data source has credentials in Aspose.Cells | throw error when external connection has no credentials Aspose.Cells | validate external connections in .NET Aspose.Cells
// Developer Intent: Guarantee that every external connection in a workbook possesses a non‑empty credential setting before invoking Save.
// Use Cases: Prevent runtime failures by blocking saves of workbooks with unsecured external data connections. | Validate credentials in template workbooks that contain pre‑configured data sources. | Integrate credential checks into automated reporting pipelines that pull data from external systems.
// AI Prompts: Create a C# method that scans workbook.DataConnections and throws InvalidOperationException if any connection's Credentials equals CredentialsMethodType.None. | Show how to log the names of connections lacking credentials before raising the exception in Aspose.Cells. | Extend the validation to accept only Integrated or Stored credentials and reject all other types.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionValidation
{
    // Shows how to loop through a Workbook’s DataConnections in C#, confirm each ExternalConnection has a non‑empty Credentials value, raise an InvalidOperationException for any missing credentials, and save the workbook only after validation succeeds.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add a dummy external connection for demonstration purposes
            // In real scenarios the workbook would already contain external connections
            // ExternalConnection connection = workbook.DataConnections.Add("MyConnection", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=sample.xlsx;", true);
            // connection.Credentials = CredentialsMethodType.Integrated; // set a non‑empty credential

            // Validate that each external connection has non‑empty credentials before saving
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                // The Credentials property is an enum; assume CredentialsMethodType.None indicates no credentials
                if (conn.Credentials == CredentialsMethodType.None)
                {
                    throw new InvalidOperationException(
                        $"External connection \"{conn.Name}\" does not have credentials set.");
                }
            }

            // Save the workbook after successful validation
            workbook.Save("ValidatedWorkbook.xlsx");
        }
    }
}
