using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    class RetrieveDbConnectionInfo
    {
        static void Main()
        {
            // Load an existing workbook that may contain external DB connections
            // Replace "input.xlsx" with the path to your workbook file
            Workbook workbook = new Workbook("input.xlsx");

            // Access the collection of external connections in the workbook
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through each connection and process only DBConnection instances
            foreach (ExternalConnection conn in connections)
            {
                if (conn is DBConnection dbConn)
                {
                    // Retrieve the command text (SQL, table name, etc.)
                    string commandText = dbConn.Command;

                    // Retrieve the command type (e.g., SqlStatement, TableName, etc.)
                    OLEDBCommandType commandType = dbConn.CommandType;

                    // Retrieve the (obsolete) connection information string
                    // Note: ConnectionInfo is marked obsolete; ConnectionString is the preferred property
                    string connectionInfo = dbConn.ConnectionInfo;

                    // Output the retrieved values for inspection
                    Console.WriteLine("DBConnection found:");
                    Console.WriteLine($"  Command Text   : {commandText}");
                    Console.WriteLine($"  Command Type   : {commandType}");
                    Console.WriteLine($"  ConnectionInfo : {connectionInfo}");
                    Console.WriteLine();
                }
            }

            // Optionally, save the workbook if any modifications were made
            // Here we simply re‑save the original file (or specify a new file name)
            workbook.Save("output.xlsx");
        }
    }
}