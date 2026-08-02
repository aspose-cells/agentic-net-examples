// Title: Set Integrated Security for a Query Table Connection in Aspose.Cells (.NET)
// Description: Loads a workbook, gets the first worksheet’s query table, updates its ExternalConnection.ConnectionString to use Windows Integrated Security (e.g., SQLOLEDB with Integrated Security=SSPI), optionally sets CredentialsMethodType to Integrated, and saves the file.
// Keywords: Aspose.Cells query table connection string | integrated security Aspose.Cells | Windows authentication Excel query table | ExternalConnection CredentialsMethodType | C# modify query table connection
// Common Searches: Aspose.Cells set query table to use integrated security | How to enable Windows authentication for a query table in .NET | Change external connection string for Excel query table programmatically | Set CredentialsMethodType to Integrated in Aspose.Cells | Secure query table connection string Aspose.Cells
// Developer Intent: Modify a query table’s external connection so it authenticates with Windows Integrated Security instead of stored credentials.
// Use Cases: Replace hard‑coded SQL credentials with Windows authentication for compliance‑driven reporting. | Automate workbook preparation before deployment to ensure all query tables use secure authentication. | Batch‑process multiple Excel files to enforce Integrated Security on their data connections.
// AI Prompts: Write C# code using Aspose.Cells to change a query table’s connection string to Windows Integrated Security and save the workbook. | Explain how to read and set the CredentialsMethodType property of an ExternalConnection for a query table. | Show how to verify that a query table’s connection string was updated after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads a workbook, gets the first worksheet’s query table, updates its ExternalConnection.ConnectionString to use Windows Integrated Security (e.g., SQLOLEDB with Integrated Security=SSPI), optionally sets CredentialsMethodType to Integrated, and saves the file.
class SetQueryTableConnectionString
{
    static void Main()
    {
        // Load an existing workbook that contains a query table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one query table
        if (sheet.QueryTables.Count > 0)
        {
            // Get the first query table
            QueryTable queryTable = sheet.QueryTables[0];

            // Retrieve the external connection associated with the query table
            ExternalConnection connection = queryTable.ExternalConnection;

            // Set the connection string to use Integrated Security (Windows authentication)
            // Example for SQL Server OLE DB provider
            connection.ConnectionString = 
                "Provider=SQLOLEDB;Data Source=MyServer;Initial Catalog=MyDatabase;Integrated Security=SSPI;";

            // Optionally, you can also set the authentication method explicitly
            connection.CredentialsMethodType = CredentialsMethodType.Integrated;
        }
        else
        {
            Console.WriteLine("No query tables found in the worksheet.");
        }

        // Save the workbook with the updated connection string
        workbook.Save("output.xlsx");
    }
}
