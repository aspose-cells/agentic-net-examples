using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class UpdateDbConnectionDescription
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing external data connections
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

                // Update the connection description
                dbConn.ConnectionDescription = "Migrated connection: now points to the new data warehouse";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"DBConnection description updated and workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Program entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateDbConnectionDescription.Run();
        }
    }
}