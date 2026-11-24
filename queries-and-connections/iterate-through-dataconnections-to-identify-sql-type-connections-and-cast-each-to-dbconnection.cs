// Title: Enumerate workbook DataConnections and extract SQL connection strings using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells iterate DataConnections to get SQL connection strings | how to filter Excel DataConnections by Sql type with Aspose.Cells | retrieve connection string from Sql DataConnection in a workbook using Aspose.Cells | using reflection to access ConnectionString property of Aspose.Cells DataConnection
// Tags: enumerate DataConnections Aspose.Cells | filter SQL DataConnection C# | extract connection string Aspose.Cells | workbook DataConnection reflection | Aspose.Cells DBConnection casting

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel workbook with Aspose.Cells, iterates through its DataConnections via reflection, identifies connections where ConnectionType equals "Sql", reads their ConnectionString values, outputs the strings, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all data connections using reflection to avoid direct type dependencies
                foreach (object connObj in workbook.DataConnections)
                {
                    try
                    {
                        Type connType = connObj.GetType();

                        // Get the ConnectionType property
                        var connTypeProp = connType.GetProperty("ConnectionType");
                        if (connTypeProp == null) continue;

                        var connTypeValue = connTypeProp.GetValue(connObj);
                        if (connTypeValue == null) continue;

                        // Check if the connection type is SQL (SQL Server)
                        if (connTypeValue.ToString() == "Sql")
                        {
                            // Retrieve the connection string
                            var connStrProp = connType.GetProperty("ConnectionString");
                            if (connStrProp != null)
                            {
                                string connectionString = connStrProp.GetValue(connObj) as string;
                                Console.WriteLine("SQL Connection String: " + connectionString);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle any unexpected errors for a single connection
                        Console.WriteLine($"Error processing a data connection: {ex.Message}");
                    }
                }

                // Save the workbook (if any modifications were made)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // General exception handling for loading/saving the workbook
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
