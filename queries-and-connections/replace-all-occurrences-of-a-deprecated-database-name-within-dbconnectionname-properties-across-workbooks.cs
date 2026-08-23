// Title: How to replace a deprecated database name in DBConnection.Name for all external data connections in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, iterates through workbook.DataConnections, and substitutes a given deprecated database name with a new one in each DBConnection.Name property. | Create a reusable C# method that takes inputPath, outputPath, oldDbName, and newDbName, then updates all DBConnection.Name values in the workbook's external connections and saves the file.
// Common Searches: Aspose.Cells replace old database name in external connections C# | C# update DBConnection Name property for all data connections in Excel file | How to rename deprecated DB name in Excel workbook using Aspose.Cells | Programmatically change database identifier in workbook data connections .NET
// Tags: replace DBConnection Name Aspose.Cells C# | update external data connections Excel Aspose.Cells | modify workbook data connections .NET | bulk rename database identifier in Excel workbook | Aspose.Cells change DBConnection property programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook with Aspose.Cells, accesses its DataConnections collection, iterates over each DBConnection, substitutes any occurrence of a deprecated database name in the Name property with a new name, and saves the updated workbook to a specified location.
    public class ReplaceDeprecatedDbNameInConnections
    {
        /// <param name="inputFile">Path to the source workbook.</param>
        /// <param name="outputFile">Path where the updated workbook will be saved.</param>
        /// <param name="oldDbName">Deprecated database name to be replaced.</param>
        /// <param name="newDbName">New database name to substitute.</param>
        public static void Run(string inputFile, string outputFile, string oldDbName, string newDbName)
        {
            try
            {
                // Verify that the input workbook exists.
                if (!File.Exists(inputFile))
                {
                    Console.Error.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputFile);

                // Access the collection of external data connections.
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Iterate through each connection.
                foreach (ExternalConnection connection in connections)
                {
                    // Process only DBConnection instances.
                    if (connection is DBConnection dbConn)
                    {
                        // Ensure the Name property is not null before replacement.
                        if (!string.IsNullOrEmpty(dbConn.Name))
                        {
                            // Replace all occurrences of the deprecated name with the new name.
                            dbConn.Name = dbConn.Name.Replace(oldDbName, newDbName);
                        }
                    }
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the desired output location.
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        // Entry point for the console application.
        private static void Main(string[] args)
        {
            // Expected arguments: inputFile outputFile oldDbName newDbName
            if (args.Length != 4)
            {
                Console.WriteLine("Usage: ReplaceDeprecatedDbNameInConnections <inputFile> <outputFile> <oldDbName> <newDbName>");
                return;
            }

            string inputFile = args[0];
            string outputFile = args[1];
            string oldDbName = args[2];
            string newDbName = args[3];

            ReplaceDeprecatedDbNameInConnections.Run(inputFile, outputFile, oldDbName, newDbName);
        }
    }
}
