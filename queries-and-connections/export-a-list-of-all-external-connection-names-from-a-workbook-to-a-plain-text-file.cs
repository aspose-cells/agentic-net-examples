using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsDemo
{
    public class ExportExternalConnectionNames
    {
        // Exports all external connection names from the specified workbook to a plain text file.
        public static void Run(string workbookPath, string outputTextFile)
        {
            try
            {
                // Verify workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.Error.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Retrieve external connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputTextFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write each connection name to the output file
                using (StreamWriter writer = new StreamWriter(outputTextFile))
                {
                    for (int i = 0; i < connections.Count; i++)
                    {
                        writer.WriteLine(connections[i].Name);
                    }
                }

                Console.WriteLine($"Exported {connections.Count} connection name(s) to {outputTextFile}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        // Entry point required for compilation
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsDemo <workbookPath> <outputTextFile>");
                return;
            }

            ExportExternalConnectionNames.Run(args[0], args[1]);
        }
    }
}