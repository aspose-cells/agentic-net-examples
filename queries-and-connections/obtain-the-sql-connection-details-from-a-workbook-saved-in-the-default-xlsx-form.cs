using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RetrieveSqlConnectionDetails
{
    static void Main()
    {
        // Path to the existing XLSX workbook
        string workbookPath = "input.xlsx";

        // Load the workbook using the constructor that accepts a file name
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through all external data connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Check if the connection is a database (SQL) connection
            if (connection is DBConnection dbConnection)
            {
                // Output the connection name and its connection string (SQL details)
                Console.WriteLine("Connection Name: " + dbConnection.Name);
                Console.WriteLine("Connection String: " + dbConnection.ConnectionString);
            }
        }

        // No modifications are made, so saving is optional.
        // If you need to save the workbook after any changes, uncomment the line below:
        // workbook.Save("output.xlsx");
    }
}