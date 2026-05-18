using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsDbConnectionInspection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook that may contain DB connections.
            // Replace the path with your actual workbook file.
            string workbookPath = "input.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // Get the collection of external data connections.
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through each connection and inspect DBConnection properties.
            foreach (ExternalConnection conn in connections)
            {
                // Check if the connection is a DBConnection.
                if (conn is DBConnection dbConn)
                {
                    // Retrieve and display the command text (Command property).
                    Console.WriteLine("Command Text: " + dbConn.Command);

                    // Retrieve and display the command type (CommandType property).
                    Console.WriteLine("Command Type: " + dbConn.CommandType);

                    // Retrieve and display the connection information string (ConnectionInfo property).
                    // Note: ConnectionInfo is obsolete; ConnectionString is the recommended property.
                    Console.WriteLine("Connection Info (obsolete): " + dbConn.ConnectionInfo);
                    Console.WriteLine("Connection String (recommended): " + dbConn.ConnectionString);

                    Console.WriteLine(new string('-', 40));
                }
            }

            // Optionally, save the workbook if any modifications were made.
            // workbook.Save("output.xlsx");
        }
    }
}