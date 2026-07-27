using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Expected arguments: inputPath outputPath oldDbName newDbName
            if (args.Length != 4)
            {
                Console.WriteLine("Usage: <inputPath> <outputPath> <oldDbName> <newDbName>");
                return;
            }

            try
            {
                ReplaceDeprecatedDbNameInConnections.Run(
                    args[0], args[1], args[2], args[3]);

                Console.WriteLine("Workbook processed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class ReplaceDeprecatedDbNameInConnections
    {
        // Replace occurrences of a deprecated database name in DBConnection properties.
        public static void Run(string inputPath, string outputPath, string oldDbName, string newDbName)
        {
            // Ensure the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Access the collection of external data connections.
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Iterate through each connection and process only DBConnection instances.
                foreach (ExternalConnection connection in connections)
                {
                    if (connection is DBConnection dbConn)
                    {
                        // Replace in the connection's Name property.
                        if (!string.IsNullOrEmpty(dbConn.Name))
                            dbConn.Name = dbConn.Name.Replace(oldDbName, newDbName);

                        // Replace in SourceFile if present.
                        if (!string.IsNullOrEmpty(dbConn.SourceFile))
                            dbConn.SourceFile = dbConn.SourceFile.Replace(oldDbName, newDbName);

                        // Replace in ConnectionString (new API, replaces obsolete ConnectionInfo).
                        if (!string.IsNullOrEmpty(dbConn.ConnectionString))
                            dbConn.ConnectionString = dbConn.ConnectionString.Replace(oldDbName, newDbName);

                        // Replace in Command.
                        if (!string.IsNullOrEmpty(dbConn.Command))
                            dbConn.Command = dbConn.Command.Replace(oldDbName, newDbName);

                        // Replace in SecondCommand (new API, replaces obsolete SeverCommand).
                        if (!string.IsNullOrEmpty(dbConn.SecondCommand))
                            dbConn.SecondCommand = dbConn.SecondCommand.Replace(oldDbName, newDbName);
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
            }
            catch
            {
                // Rethrow to allow the caller to handle the exception.
                throw;
            }
        }
    }
}