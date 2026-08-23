// Title: Export all external data connection names from an Excel workbook to a text file using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console program that opens an .xlsx file with Aspose.Cells, iterates over workbook.DataConnections, and writes each connection's Name to a .txt file. | Generate C# code that verifies the source workbook exists, creates the output folder if necessary, and saves the list of external connection names to a plain‑text file while handling exceptions. | Provide a C# snippet using Aspose.Cells to retrieve the Name property of every DataConnection in a workbook and export those names to a text document.
// Common Searches: how to list external data connections in an Excel file using Aspose.Cells C# | C# Aspose.Cells export workbook DataConnections names to txt | save Excel external connection names to a text file with Aspose.Cells | retrieve DataConnection.Name collection Aspose.Cells .NET example
// Tags: export DataConnections names to text file Aspose.Cells | Aspose.Cells list external connections C# | write Excel connection names plain text .NET | enumerate workbook DataConnections Aspose.Cells | save external connection list txt Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalConnectionsExport
{
    // The C# console application loads a specified .xlsx workbook with Aspose.Cells, checks for the file's existence, creates the output directory if needed, iterates through the workbook's DataConnections collection, writes each connection's Name to a plain‑text file, and reports the number of exported connections while handling errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Path to the output text file
            string outputPath = "ExternalConnectionNames.txt";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Input file '{workbookPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Use dynamic to avoid compile‑time dependency on DataConnection types
                dynamic connections = workbook.DataConnections;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                int exportedCount = 0;

                // Write each connection name to the text file
                using (StreamWriter writer = new StreamWriter(outputPath, false))
                {
                    foreach (var connObj in connections)
                    {
                        try
                        {
                            dynamic conn = connObj;
                            writer.WriteLine(conn.Name);
                            exportedCount++;
                        }
                        catch (Exception innerEx)
                        {
                            Console.WriteLine($"Failed to process a connection: {innerEx.Message}");
                        }
                    }
                }

                Console.WriteLine($"Exported {exportedCount} connection name(s) to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
