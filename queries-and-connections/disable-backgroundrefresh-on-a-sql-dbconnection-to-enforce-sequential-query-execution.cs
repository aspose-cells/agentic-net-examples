// Title: Disable BackgroundRefresh on a SQL DBConnection in Aspose.Cells for .NET to enforce sequential query execution
// Common Searches: Aspose.Cells C# set DBConnection.BackgroundRefresh to false | make external SQL connection refresh synchronously in Aspose.Cells workbook | prevent background data refresh for DBConnection in .NET Aspose.Cells | how to enforce sequential query execution with Aspose.Cells DBConnection | disable asynchronous refresh of SQL DBConnection using Aspose.Cells API
// Tags: DBConnection BackgroundRefresh property Aspose.Cells | synchronous SQL data refresh Aspose.Cells .NET | external DB connection configuration Aspose.Cells | sequential query execution workbook Aspose.Cells | disable async DBConnection refresh C#

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace DisableBackgroundRefreshDemo
{
    // The example creates a new Workbook, adds a DBConnection to its DataConnections collection, configures the SQL connection string and command, sets BackgroundRefresh to false to enforce synchronous query execution, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "DBConnection_BackgroundRefreshDisabled.xlsx";

            try
            {
                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Get the collection of external data connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Create a new DBConnection instance by accessing the next index in the collection
                // This automatically adds a new DBConnection to the workbook
                DBConnection dbConnection = (DBConnection)connections[connections.Count];

                // Configure the DB connection – replace with your actual connection details
                dbConnection.ConnectionString = "Provider=SQLOLEDB;Data Source=YOUR_SERVER;Initial Catalog=YOUR_DATABASE;Integrated Security=SSPI;";
                dbConnection.CommandType = OLEDBCommandType.SqlStatement;
                dbConnection.Command = "SELECT TOP 10 * FROM dbo.YourTable";

                // Disable background refresh to enforce synchronous execution
                dbConnection.BackgroundRefresh = false;

                // Verify settings
                Console.WriteLine($"BackgroundRefresh is set to: {dbConnection.BackgroundRefresh}");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
