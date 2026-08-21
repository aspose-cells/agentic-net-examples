// Title: Read CommandText, CommandType, and ConnectionInfo from DBConnection objects in Aspose.Cells (C#)
// Description: Loads an Excel workbook, iterates its DataConnections, filters for DBConnection instances, and extracts the SQL command text, OLEDB command type, and legacy connection string, outputting the values to the console.
// Keywords: Aspose.Cells | DBConnection | CommandText | CommandType | ConnectionInfo | external data connections | C# | .NET | Excel workbook | retrieve DB connection details
// Common Searches: Aspose.Cells get DBConnection command text | How to read CommandType from Excel external connection using Aspose.Cells | Retrieve ConnectionInfo from DBConnection in C# | Iterate DataConnections Aspose.Cells .NET | Extract DB connection properties from workbook
// Developer Intent: The developer needs to enumerate DBConnection objects in a workbook and display each connection's command text, command type, and connection information.
// Use Cases: Audit embedded SQL queries in an Excel file for compliance. | Validate that external data connections use the correct command type before data import. | Generate a report of all database sources referenced by a workbook.
// AI Prompts: Write C# code with Aspose.Cells that collects Command, CommandType, and ConnectionInfo from every DBConnection and stores them in a list. | Explain safe handling of the obsolete ConnectionInfo property and recommend a modern alternative. | Show how to filter DataConnections to DBConnection only and export their details to a CSV file.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads an Excel workbook, iterates its DataConnections, filters for DBConnection instances, and extracts the SQL command text, OLEDB command type, and legacy connection string, outputting the values to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook that may contain DB connections
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external connections
        foreach (ExternalConnection conn in workbook.DataConnections)
        {
            // Process only DBConnection objects
            if (conn is DBConnection dbConn)
            {
                // Retrieve the command text, command type, and connection info
                string commandText = dbConn.Command;
                OLEDBCommandType commandType = dbConn.CommandType;
                string connectionInfo = dbConn.ConnectionInfo; // Obsolete property, still accessible

                // Display the retrieved information
                Console.WriteLine("DBConnection found:");
                Console.WriteLine($"  CommandText   : {commandText}");
                Console.WriteLine($"  CommandType   : {commandType}");
                Console.WriteLine($"  ConnectionInfo: {connectionInfo}");
                Console.WriteLine();
            }
        }

        // Save the workbook (no modifications made, but required by lifecycle rules)
        workbook.Save("output.xlsx");
    }
}
