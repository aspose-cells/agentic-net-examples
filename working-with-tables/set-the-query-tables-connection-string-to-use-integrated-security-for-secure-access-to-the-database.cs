using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Load a workbook that already contains a query table.
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed).
        Worksheet sheet = workbook.Worksheets[0];

        // Verify that the worksheet has at least one query table.
        if (sheet.QueryTables.Count > 0)
        {
            // Retrieve the first query table.
            QueryTable queryTable = sheet.QueryTables[0];

            // Obtain the connection id associated with the query table.
            int connectionId = queryTable.ConnectionId;

            // Get the external connection object from the workbook's collection.
            ExternalConnection externalConn = workbook.DataConnections[connectionId] as ExternalConnection;

            // Ensure the connection is a DBConnection (OLE DB/ODBC).
            if (externalConn is DBConnection dbConn)
            {
                // Set the connection string to use Integrated Security (Windows authentication).
                dbConn.ConnectionString = "Provider=SQLOLEDB;Data Source=MyServer;Initial Catalog=MyDatabase;Integrated Security=SSPI;";

                // Optionally specify the command and its type.
                dbConn.CommandType = OLEDBCommandType.SqlStatement;
                dbConn.Command = "SELECT * FROM MyTable";
            }
        }

        // Save the workbook with the updated connection string.
        workbook.Save("output.xlsx");
    }
}