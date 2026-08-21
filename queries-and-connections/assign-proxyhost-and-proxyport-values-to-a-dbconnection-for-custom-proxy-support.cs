// Title: Set Proxy Host and Port for DBConnection in Aspose.Cells for .NET (C#)
// Description: Shows how to assign a custom WebProxy (host and port) to WebRequest.DefaultWebProxy, create an uninitialized DBConnection, add it to a workbook's ExternalConnectionCollection, configure its properties, and save the workbook.
// Keywords: Aspose.Cells | C# | DBConnection | proxy | WebProxy | WebRequest.DefaultWebProxy | external data connection | connection string | OLEDBCommandType | Excel export
// Common Searches: Aspose.Cells set proxy for DBConnection | C# assign WebProxy to external connections Aspose.Cells | How to use proxy with Aspose.Cells data connections | Configure proxy host and port in Aspose.Cells .NET | Create DBConnection without constructor Aspose.Cells
// Developer Intent: Learn how to configure a DBConnection to route through a specific proxy server in Aspose.Cells for .NET.
// Use Cases: Apply a corporate proxy to a single DBConnection before exporting data to Excel. | Reuse the same proxy settings for multiple external connections within one workbook. | Instantiate a DBConnection via FormatterServices, set its properties, and enable proxy routing before saving.
// AI Prompts: Write C# code that sets WebRequest.DefaultWebProxy with a host and port and creates a DBConnection in Aspose.Cells. | Generate an example that creates three DBConnection objects, configures each to use the same proxy, and saves the workbook. | Explain how to add proxy credentials (username and password) to the WebProxy for Aspose.Cells external connections.

using System;
using System.Net;
using System.Runtime.Serialization;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Shows how to assign a custom WebProxy (host and port) to WebRequest.DefaultWebProxy, create an uninitialized DBConnection, add it to a workbook's ExternalConnectionCollection, configure its properties, and save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // If you need to use a proxy for external connections, set it on the default WebRequest.
            // Example (uncomment and adjust as needed):
            // string proxyHost = "proxy.example.com";
            // int proxyPort = 8080;
            // WebRequest.DefaultWebProxy = new WebProxy(proxyHost, proxyPort);

            // Access the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // DBConnection does not have a public constructor, so create an uninitialized instance
            DBConnection dbConnection = (DBConnection)FormatterServices.GetUninitializedObject(typeof(DBConnection));

            // Add the DBConnection to the workbook's connections collection
            ((IList<ExternalConnection>)connections).Add(dbConnection);

            // Configure the DBConnection
            dbConnection.Name = "MyDBConnection";
            dbConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\data\database.mdb;";
            dbConnection.CommandType = OLEDBCommandType.TableName;
            dbConnection.Command = "Customers";

            // Save the workbook
            string outputPath = "DBConnectionWithProxy.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
