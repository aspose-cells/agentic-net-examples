using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class IdentifySqlDbConnections
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook that may contain external data connections
                Workbook workbook = new Workbook(inputPath);

                // Get the collection of external connections from the workbook
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Iterate through each connection in the collection
                for (int i = 0; i < connections.Count; i++)
                {
                    ExternalConnection conn = connections[i];

                    // Process only DB connections (ODBC or OLE DB)
                    if (conn is DBConnection dbConn)
                    {
                        // Determine if the DBConnection uses a SQL statement
                        bool isSql = dbConn.CommandType == OLEDBCommandType.SqlStatement;

                        // Output information about the connection
                        Console.WriteLine($"Connection #{i + 1}:");
                        Console.WriteLine($"  Name          : {dbConn.Name}");
                        Console.WriteLine($"  ClassType     : {dbConn.ClassType}");
                        Console.WriteLine($"  SourceType    : {dbConn.SourceType}");
                        Console.WriteLine($"  CommandType   : {dbConn.CommandType}");
                        Console.WriteLine($"  Is SQL Type   : {isSql}");
                        Console.WriteLine($"  Command       : {dbConn.Command}");
                        Console.WriteLine();
                    }
                }

                // Save the workbook (even if unchanged) to illustrate a complete flow
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}