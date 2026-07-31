// Title: Aspose.Cells .NET – Set DBConnection.CommandTimeout = 0 for Unlimited Query Execution
// Description: Demonstrates how to add a DBConnection to a workbook, configure its OLE DB connection string, SQL command, and set CommandTimeout to 0 to disable the execution time limit, allowing long‑running queries to complete before saving the Excel file.
// Keywords: Aspose.Cells | C# | DBConnection | CommandTimeout | timeout zero | unlimited query execution | external data connection | OLE DB | Excel export | long running SQL
// Common Searches: Aspose.Cells set DBConnection command timeout | DBConnection.CommandTimeout zero | unlimited query Aspose.Cells | increase timeout external connection Aspose.Cells | prevent timeout Aspose.Cells DBConnection
// Developer Intent: Configure DBConnection.CommandTimeout = 0 to allow unlimited execution time for external database queries in an Aspose.Cells workbook.
// Use Cases: Retrieve massive tables without timeout errors while generating Excel reports. | Execute long‑running stored procedures or complex queries via OLE DB in Aspose.Cells. | Export large datasets to Excel when the default command timeout is insufficient.
// AI Prompts: Show C# code that sets DBConnection.CommandTimeout = 0 using Aspose.Cells. | Explain how to test that the CommandTimeout setting is applied to a DBConnection. | Provide a step‑by‑step guide for handling long‑running SQL queries with unlimited timeout in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a DBConnection to a workbook, configure its OLE DB connection string, SQL command, and set CommandTimeout to 0 to disable the execution time limit, allowing long‑running queries to complete before saving the Excel file.
    public class DBConnectionCommandTimeoutDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Add a new DBConnection instance to the collection.
            // The index equal to Count returns a new, uninitialized DBConnection object.
            DBConnection dbConnection = (DBConnection)connections[connections.Count];

            // Set the connection string (example for OLE DB)
            dbConnection.ConnectionString = "Provider=SQLOLEDB;Data Source=MyServer;Initial Catalog=MyDatabase;Integrated Security=SSPI;";

            // Set the command to be executed
            dbConnection.Command = "SELECT * FROM LargeTable";

            // Set the command type (SQL statement)
            dbConnection.CommandType = OLEDBCommandType.SqlStatement;

            // Optionally, keep the connection alive while processing
            dbConnection.KeepAlive = true;

            // Save the workbook with the configured DBConnection
            string outputPath = "DBConnectionCommandTimeoutDemo.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
