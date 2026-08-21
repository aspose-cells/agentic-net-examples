// Title: Export Excel External Connection Names to Text with Aspose.Cells for .NET
// Description: Load an Excel workbook using Aspose.Cells, read its DataConnections collection, and write each external connection name to a plain‑text file (one per line). Includes checks for missing files and robust error handling.
// Keywords: Aspose.Cells | C# | .NET | DataConnections | external connection names | export to text file | Excel workbook | sample code | API example | list connections
// Common Searches: Aspose.Cells list external connections | C# export Excel data connection names to txt | how to get connection names from workbook using Aspose.Cells | save Excel data connections as plain text .NET | extract external data source names from Excel with Aspose
// Developer Intent: Retrieve all external connection names from an Excel workbook and write them to a plain‑text file.
// Use Cases: Audit all external data sources referenced in a workbook before migration. | Create an inventory file for scripts that need to validate or rename connections. | Generate a quick reference list for documentation or compliance reviews.
// AI Prompts: Generate C# code with Aspose.Cells that writes each external connection name from a workbook to a CSV file. | Show how to handle a workbook that contains no external connections when exporting names using Aspose.Cells. | Provide an example that logs the index and name of each exported connection while writing to a text file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalConnectionsExport
{
    // Load an Excel workbook using Aspose.Cells, read its DataConnections collection, and write each external connection name to a plain‑text file (one per line). Includes checks for missing files and robust error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook
            string workbookPath = "InputWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Get the collection of external connections from the workbook
                var connections = workbook.DataConnections;

                // Path for the output plain text file
                string outputPath = "ExternalConnectionNames.txt";

                // Write each connection name to the text file, one per line
                using (StreamWriter writer = new StreamWriter(outputPath))
                {
                    for (int i = 0; i < connections.Count; i++)
                    {
                        // Retrieve the name of the current external connection
                        string connectionName = connections[i].Name;

                        // Write the name to the file
                        writer.WriteLine(connectionName);
                    }
                }

                // Inform the user that the operation completed
                Console.WriteLine($"Exported {connections.Count} external connection name(s) to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
