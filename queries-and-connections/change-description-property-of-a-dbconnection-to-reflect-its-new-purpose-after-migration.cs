using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class UpdateDbConnectionDescription
    {
        // Entry point required for the console application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains a DBConnection
                Workbook workbook = new Workbook(inputPath);

                // Access the collection of external connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Find the first DBConnection in the collection
                DBConnection dbConn = null;
                foreach (ExternalConnection conn in connections)
                {
                    if (conn is DBConnection dbConnection)
                    {
                        dbConn = dbConnection;
                        break;
                    }
                }

                if (dbConn == null)
                {
                    Console.WriteLine("No DBConnection objects found in the workbook.");
                    return;
                }

                // Update the description to reflect the new purpose after migration
                dbConn.ConnectionDescription = "Migrated connection: now points to the new data warehouse";

                // Optionally, update other properties such as the connection string
                // dbConn.ConnectionString = "Provider=SQLNCLI11;Data Source=NewServer;Initial Catalog=NewDB;Integrated Security=SSPI;";

                // Save the workbook with the modified connection
                workbook.Save(outputPath);

                Console.WriteLine("DBConnection description updated and workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}