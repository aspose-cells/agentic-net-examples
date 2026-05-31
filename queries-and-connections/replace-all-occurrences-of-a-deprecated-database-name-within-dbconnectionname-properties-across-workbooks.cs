using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class ReplaceDbConnectionName
    {
        /// <summary>
        /// Replaces occurrences of a deprecated database name within DBConnection.Name properties.
        /// </summary>
        /// <param name="inputPath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the modified workbook will be saved.</param>
        /// <param name="oldDbName">Deprecated database name to be replaced.</param>
        /// <param name="newDbName">New database name to use.</param>
        public static void ProcessWorkbook(string inputPath, string outputPath, string oldDbName, string newDbName)
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all external connections in the workbook
            foreach (ExternalConnection connection in workbook.DataConnections)
            {
                // Check if the connection is a DBConnection
                if (connection is DBConnection dbConn)
                {
                    // Replace the deprecated name in the DBConnection.Name property
                    if (!string.IsNullOrEmpty(dbConn.Name))
                    {
                        dbConn.Name = dbConn.Name.Replace(oldDbName, newDbName);
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            string inputFile = "input.xlsx";
            string outputFile = "output.xlsx";
            string deprecatedName = "OldDatabase";
            string updatedName = "NewDatabase";

            ProcessWorkbook(inputFile, outputFile, deprecatedName, updatedName);

            Console.WriteLine($"Processed workbook saved to '{outputFile}'.");
        }
    }
}