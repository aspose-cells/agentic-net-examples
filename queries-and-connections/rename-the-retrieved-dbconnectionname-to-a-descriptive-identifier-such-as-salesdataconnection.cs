using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class RenameDbConnectionDemo
    {
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

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that contains a DBConnection
                Workbook workbook = new Workbook(inputPath);

                // Access the collection of external connections in the workbook
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Iterate through the connections to find DBConnection objects
                foreach (ExternalConnection connection in connections)
                {
                    if (connection is DBConnection dbConn)
                    {
                        // Rename the connection to a more descriptive identifier
                        dbConn.Name = "SalesDataConnection";
                        Console.WriteLine($"DBConnection renamed to: {dbConn.Name}");
                    }
                }

                // Save the workbook with the updated connection name
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}