// Title: Rename Excel DBConnection and Update Connection String with Aspose.Cells for .NET after Server Migration
// Description: Shows how to open an .xlsx file, loop through workbook.DataConnections, find DBConnection objects, replace the old server identifier in both the connection name and connection string, and save the modified workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel DBConnection rename | update connection string | external data connection | server migration | workbook.DataConnections | replace server name | Excel automation | database server change
// Common Searches: how to rename DBConnection in Excel using Aspose.Cells | update Excel external connection string after server move .NET | Aspose.Cells change database server in workbook | C# iterate workbook.DataConnections to modify DBConnection | replace old server name in Excel DBConnection
// Developer Intent: Change the name and connection string of an existing DBConnection so it points to a new database server.
// Use Cases: Load a workbook and enumerate its DataConnections to locate DBConnection entries. | Swap the old server identifier with the new one in the DBConnection.Name property. | Update DBConnection.ConnectionString to reference the new server before saving the file.
// AI Prompts: Generate C# code with Aspose.Cells that renames a DBConnection and updates its connection string after a server migration. | Explain safe iteration of workbook.DataConnections to modify only DBConnection objects. | Provide error‑handling patterns for missing Excel files and connection update failures in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Shows how to open an .xlsx file, loop through workbook.DataConnections, find DBConnection objects, replace the old server identifier in both the connection name and connection string, and save the modified workbook using Aspose.Cells for .NET.
    public class RenameDbConnectionDemo
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook that contains the existing DBConnection
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all external connections in the workbook
                foreach (ExternalConnection connection in workbook.DataConnections)
                {
                    // Process only DBConnection objects
                    if (connection is DBConnection dbConn)
                    {
                        // Replace old server identifier "OldServer" with "NewServer" in the connection name
                        if (!string.IsNullOrEmpty(dbConn.Name) && dbConn.Name.Contains("OldServer"))
                        {
                            dbConn.Name = dbConn.Name.Replace("OldServer", "NewServer");
                        }

                        // Also update the connection string to point to the new server
                        if (!string.IsNullOrEmpty(dbConn.ConnectionString) && dbConn.ConnectionString.Contains("OldServer"))
                        {
                            dbConn.ConnectionString = dbConn.ConnectionString.Replace("OldServer", "NewServer");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RenameDbConnectionDemo.Run();
        }
    }
}
