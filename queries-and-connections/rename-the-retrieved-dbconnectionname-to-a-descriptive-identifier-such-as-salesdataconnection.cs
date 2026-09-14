// Title: Rename external DBConnection to a custom name (e.g., SalesDataConnection) in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Set the Name property of every DBConnection in a loaded Workbook to a specific string and save the file with Aspose.Cells in C#. | Iterate through workbook.DataConnections, detect DBConnection objects, and assign a descriptive identifier such as 'SalesDataConnection' using the Aspose.Cells API. | Programmatically update the external database connection name in an existing .xlsx file and write the changes to a new file with Aspose.Cells for .NET.
// Common Searches: how to change the name of a DBConnection in an Excel file using Aspose.Cells C# | Aspose.Cells rename external database connection identifier in workbook | C# code to set DBConnection.Name property for all data connections in .xlsx | update Excel data connection name programmatically with Aspose.Cells .NET
// Tags: rename DBConnection Aspose.Cells C# | set DBConnection.Name property .xlsx | modify external data connection identifier Excel .NET | Aspose.Cells workbook data connections handling | update Excel DB connection name programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace RenameDbConnectionDemoApp
{
    // The example loads an existing Excel workbook, loops through its DataConnections, finds DBConnection objects, changes each connection's Name to "SalesDataConnection", and saves the modified workbook to a new file, with error handling for load and save operations.
    public class Program
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        private static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load workbook containing external DB connections
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through external connections and rename DB connections
            foreach (ExternalConnection connection in workbook.DataConnections)
            {
                if (connection is DBConnection dbConn)
                {
                    dbConn.Name = "SalesDataConnection";
                    Console.WriteLine($"Renamed DBConnection to: {dbConn.Name}");
                }
            }

            try
            {
                // Save updated workbook
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
