using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    /// <summary>
    /// Demonstrates how to replace a deprecated database name within DBConnection.Name
    /// properties for one or more workbooks.
    /// </summary>
    public class ReplaceDbConnectionName
    {
        /// <summary>
        /// Replaces occurrences of <paramref name="oldDbName"/> with <paramref name="newDbName"/>
        /// in the Name property of all DBConnection objects inside the workbook located at <paramref name="inputFile"/>.
        /// The modified workbook is saved to <paramref name="outputFile"/>.
        /// </summary>
        public static void Run(string inputFile, string outputFile, string oldDbName, string newDbName)
        {
            // Load the workbook (lifecycle rule: use provided load logic)
            Workbook workbook = new Workbook(inputFile);

            // Iterate through all external connections in the workbook
            ExternalConnectionCollection connections = workbook.DataConnections;
            foreach (ExternalConnection connection in connections)
            {
                // Process only DBConnection instances
                if (connection is DBConnection dbConn)
                {
                    // Replace the deprecated database name in the connection's Name property
                    if (!string.IsNullOrEmpty(dbConn.Name))
                    {
                        dbConn.Name = dbConn.Name.Replace(oldDbName, newDbName);
                    }

                    // Optional: also replace in other string properties if needed
                    // (shown here for completeness, not required by the task)
                    // dbConn.SourceFile = dbConn.SourceFile?.Replace(oldDbName, newDbName);
                    // dbConn.ConnectionInfo = dbConn.ConnectionInfo?.Replace(oldDbName, newDbName);
                    // dbConn.Command = dbConn.Command?.Replace(oldDbName, newDbName);
                    // if (!string.IsNullOrEmpty(dbConn.SeverCommand))
                    //     dbConn.SeverCommand = dbConn.SeverCommand.Replace(oldDbName, newDbName);
                }
            }

            // Save the modified workbook (lifecycle rule: use provided save logic)
            workbook.Save(outputFile);
        }

        // Example usage
        public static void Main()
        {
            // Define input and output workbook paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Define the deprecated database name and its replacement
            string oldDatabaseName = "OldDatabase";
            string newDatabaseName = "NewDatabase";

            // Execute the replacement
            Run(inputPath, outputPath, oldDatabaseName, newDatabaseName);

            Console.WriteLine($"Database name replacement completed. Saved to '{outputPath}'.");
        }
    }
}