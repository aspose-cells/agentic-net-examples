// Title: Replace Deprecated DBConnection Name in Excel Workbooks with Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through its DataConnections, identifies DBConnection objects, and substitutes a specified old database name with a new one in the DBConnection.Name property before saving the file.
// Keywords: Aspose.Cells | C# | DBConnection | replace connection name | external data connections | Excel workbook | .NET | update database name | bulk rename connections | data source migration
// Common Searches: how to rename DBConnection in an Excel file using Aspose.Cells | replace old database name in workbook connections .NET | update DBConnection.Name property across all connections | Aspose.Cells change external data source name | bulk edit Excel connection names programmatically
// Developer Intent: Programmatically change every DBConnection.Name in a workbook to replace a deprecated database identifier with a new one.
// Use Cases: Migrate legacy reports to a new database by updating connection names automatically. | Run a batch job that processes multiple Excel files and aligns their DBConnection names with a refreshed data source. | Integrate a validation step in CI/CD pipelines to enforce naming standards for external data connections.
// AI Prompts: Generate C# code using Aspose.Cells that scans a workbook's DataConnections and replaces a given old database name with a new one in DBConnection.Name. | Provide a robust version of the DBConnection rename routine with error handling, logging, and a summary of changed connections. | Create a script that iterates over a folder of Excel files and applies the database name replacement to each workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, iterates through its DataConnections, identifies DBConnection objects, and substitutes a specified old database name with a new one in the DBConnection.Name property before saving the file.
    public class ReplaceDeprecatedDbNameInConnections
    {
        /// <param name="inputFilePath">Path to the source workbook.</param>
        /// <param name="outputFilePath">Path where the modified workbook will be saved.</param>
        /// <param name="oldDbName">Deprecated database name to be replaced.</param>
        /// <param name="newDbName">New database name to use.</param>
        public static void Run(string inputFilePath, string outputFilePath, string oldDbName, string newDbName)
        {
            // Load the workbook (lifecycle rule: use provided load mechanism)
            Workbook workbook = new Workbook(inputFilePath);

            // Access the collection of external data connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through each connection
            foreach (ExternalConnection connection in connections)
            {
                // Process only DBConnection instances
                if (connection is DBConnection dbConn)
                {
                    // Replace the deprecated name in the connection's Name property
                    if (!string.IsNullOrEmpty(dbConn.Name) && dbConn.Name.Contains(oldDbName))
                    {
                        dbConn.Name = dbConn.Name.Replace(oldDbName, newDbName);
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: use provided save mechanism)
            workbook.Save(outputFilePath);
        }

        // Example usage
        public static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string deprecatedName = "OldDatabase";
            string updatedName = "NewDatabase";

            Run(inputPath, outputPath, deprecatedName, updatedName);
            Console.WriteLine("Database name replacement completed.");
        }
    }
}
