using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RenameDbConnection
{
    static void Main()
    {
        // Load the workbook that contains the DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Check if the connection is a DBConnection
            if (connection is DBConnection dbConnection)
            {
                // Rename the connection to a descriptive identifier
                dbConnection.Name = "SalesDataConnection";
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}