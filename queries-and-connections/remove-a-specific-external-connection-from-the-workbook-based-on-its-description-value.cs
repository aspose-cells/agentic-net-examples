// Title: Remove an external data connection by description using Aspose.Cells for .NET (C#)
// Description: The sample loads an Excel workbook, iterates the DataConnections collection in reverse, matches each connection's ConnectionDescription against a given string, deletes the matching connections, and saves the workbook to a new file.
// Keywords: Aspose.Cells remove data connection | C# delete Excel external connection | DataConnections collection Aspose | remove workbook connection by description | Aspose.Cells .NET API
// Common Searches: how to delete an external data connection with Aspose.Cells C# | remove Excel connection by description programmatically | Aspose.Cells iterate DataConnections and delete | C# code to strip specific workbook connections | Aspose.Cells remove stale ODBC connection
// Developer Intent: The developer needs to programmatically delete a specific external data connection from an Excel workbook when its description matches a supplied value.
// Use Cases: Clean up obsolete ODBC or web‑query connections before sharing a workbook. | Eliminate a SharePoint list connection after data refresh to avoid accidental exposure. | Strip sensitive connection details from a template before distributing it to partners.
// AI Prompts: Generate C# code with Aspose.Cells that removes all data connections whose description contains a given keyword. | Suggest robust error‑handling patterns for deleting connections in an Aspose.Cells workbook. | Show how to list every connection description in a workbook using Aspose.Cells before deciding which to remove.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, iterates the DataConnections collection in reverse, matches each connection's ConnectionDescription against a given string, deletes the matching connections, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define file paths
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Description of the external connection to remove
            string targetDescription = "Target Connection Description";

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Remove matching data connections safely
                try
                {
                    for (int i = workbook.DataConnections.Count - 1; i >= 0; i--)
                    {
                        // Use dynamic to avoid compile‑time dependency on DataConnection type
                        dynamic conn = workbook.DataConnections[i];
                        if (conn != null && conn.ConnectionDescription == targetDescription)
                        {
                            workbook.DataConnections.RemoveAt(i);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while processing data connections: {ex.Message}");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
