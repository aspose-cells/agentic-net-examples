using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class RenameDbConnectionDemo
    {
        // Entry point for the application
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
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that contains the DBConnection
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all external connections in the workbook
                foreach (ExternalConnection connection in workbook.DataConnections)
                {
                    // Process only DBConnection objects
                    if (connection is DBConnection dbConn)
                    {
                        // Example: replace old server identifier with new server identifier in the connection name
                        // Adjust the strings below to match your actual naming convention
                        const string oldServerIdentifier = "OldServer";
                        const string newServerIdentifier = "NewServer";

                        // Update connection name if it contains the old identifier
                        if (!string.IsNullOrEmpty(dbConn.Name) && dbConn.Name.Contains(oldServerIdentifier))
                        {
                            dbConn.Name = dbConn.Name.Replace(oldServerIdentifier, newServerIdentifier);
                        }

                        // Update connection string if it contains the old identifier
                        if (!string.IsNullOrEmpty(dbConn.ConnectionString) && dbConn.ConnectionString.Contains(oldServerIdentifier))
                        {
                            dbConn.ConnectionString = dbConn.ConnectionString.Replace(oldServerIdentifier, newServerIdentifier);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}