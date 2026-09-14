// Title: How to rename a DBConnection and update its connection string after server migration using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook, loop through Workbook.DataConnections, locate DBConnection objects, replace the old server identifier with the new one in both the connection's Name and ConnectionString, and save the workbook. | Programmatically adjust external database connections in an .xlsx file to reflect a changed database server using the Aspose.Cells C# API.
// Common Searches: C# Aspose.Cells replace old SQL server name in external DBConnection of an Excel file | Update DBConnection connection string after migrating database server with Aspose.Cells .NET | How to change the server reference in Excel data connections using Aspose.Cells API
// Tags: rename DBConnection server Aspose.Cells | update external connection string Excel .NET | modify workbook data connections C# | Aspose.Cells change DBConnection name | Excel file server migration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // The example loads an existing Excel workbook, iterates through its DataConnections, identifies DBConnection objects, replaces occurrences of an old server name with a new one in both the connection's Name and ConnectionString, and saves the updated workbook.
    public class RenameDbConnectionDemo
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook that contains the existing DBConnection
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all external connections in the workbook
                foreach (ExternalConnection connection in workbook.DataConnections)
                {
                    // Process only DBConnection objects
                    if (connection is DBConnection dbConn)
                    {
                        const string oldServer = "OldServer";
                        const string newServer = "NewServer";

                        // Rename the connection if it contains the old server identifier
                        if (!string.IsNullOrEmpty(dbConn.Name) && dbConn.Name.Contains(oldServer))
                        {
                            dbConn.Name = dbConn.Name.Replace(oldServer, newServer);
                        }

                        // Update the connection string if it contains the old server name
                        if (!string.IsNullOrEmpty(dbConn.ConnectionString) && dbConn.ConnectionString.Contains(oldServer))
                        {
                            dbConn.ConnectionString = dbConn.ConnectionString.Replace(oldServer, newServer);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
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
