// Title: Read Query Table Metadata (Connection String & Command Type) with Aspose.Cells for .NET
// Description: C# example that loads an Excel workbook, checks the first worksheet for query tables, and extracts each table's name, external connection string, and command type. The sample also notes that refresh‑interval settings are not exposed in the current API and saves the workbook after reading the metadata.
// Keywords: Aspose.Cells query table metadata | C# read connection string Excel | external connection command type Aspose.Cells | Aspose.Cells .NET query table example | Excel query table refresh interval limitation
// Common Searches: how to get connection string of a query table using Aspose.Cells .NET | retrieve command type from Excel query table with Aspose.Cells | check if worksheet contains query tables in C# | Aspose.Cells API for query table metadata | get refresh interval of query table Aspose.Cells
// Developer Intent: Obtain the external connection string and command type of a query table inside an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Load a workbook and programmatically list the first query table's name, connection string, and command type. | Validate the presence of query tables before accessing their metadata to avoid runtime errors. | Demonstrate a complete workflow by saving the workbook after metadata extraction.
// AI Prompts: Generate C# code with Aspose.Cells that enumerates all query tables in a workbook and prints each table's name, connection string, and command type. | Explain how to safely handle query tables that lack an external connection when reading metadata with Aspose.Cells for .NET. | Provide guidance on retrieving or approximating refresh‑interval settings for query tables, given the current API constraints.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace ReadQueryTableMetadataApp
{
    // C# example that loads an Excel workbook, checks the first worksheet for query tables, and extracts each table's name, external connection string, and command type. The sample also notes that refresh‑interval settings are not exposed in the current API and saves the workbook after reading the metadata.
    class ReadQueryTableMetadata
    {
        static void Main()
        {
            try
            {
                string inputPath = "InputWithQueryTable.xlsx";

                // Verify that the input file exists before attempting to load it.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook that contains a query table.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one query table.
                if (worksheet.QueryTables.Count > 0)
                {
                    // Get the first query table.
                    QueryTable queryTable = worksheet.QueryTables[0];
                    Console.WriteLine("Query Table Name: " + queryTable.Name);

                    // Retrieve the associated external connection.
                    ExternalConnection externalConnection = queryTable.ExternalConnection;

                    if (externalConnection != null)
                    {
                        // Connection string used to connect to the external data source.
                        Console.WriteLine("Connection String: " + externalConnection.ConnectionString);

                        // Command type (e.g., SQL, Table, etc.).
                        Console.WriteLine("Command Type: " + externalConnection.CommandType);
                    }
                    else
                    {
                        Console.WriteLine("No external connection associated with this query table.");
                    }

                    // Note: Refresh settings are not directly exposed in the current API version.
                }
                else
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                }

                // Save the workbook if any modifications were made (optional).
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
