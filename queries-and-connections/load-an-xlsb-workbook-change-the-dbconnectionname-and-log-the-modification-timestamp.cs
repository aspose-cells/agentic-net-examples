// Title: Rename the first data connection in an XLSB workbook and log the change timestamp using Aspose.Cells for .NET
// Common Searches: how to change the name of a data connection in an XLSB file with Aspose.Cells C# | Aspose.Cells C# example to rename first DBConnection and output modification time | load an XLSB workbook, update data connection name, and save with Aspose.Cells | C# code to check data connections count in an XLSB using Aspose.Cells | log timestamp after modifying a data connection name in an Aspose.Cells workbook
// Tags: rename data connection Aspose.Cells | modify DBConnection.Name C# | XLSB workbook data connections Aspose.Cells | log modification timestamp C# Aspose.Cells | save workbook as XLSB Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample verifies the input XLSB file, loads it with Aspose.Cells, checks for existing data connections, renames the first connection to "NewConnectionName", writes the change timestamp to the console, and saves the workbook back to XLSB format, handling any exceptions that may occur.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsb";
            const string outputPath = "output.xlsb";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the XLSB workbook
                Workbook workbook = new Workbook(inputPath);

                // Check for any data connections in the workbook
                if (workbook.DataConnections.Count > 0)
                {
                    // Modify the first connection's name
                    var connection = workbook.DataConnections[0];
                    connection.Name = "NewConnectionName";

                    // Log the modification timestamp
                    Console.WriteLine($"DataConnection name changed at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                }
                else
                {
                    Console.WriteLine("No data connection found in the workbook.");
                }

                // Save the workbook back to XLSB format
                workbook.Save(outputPath, SaveFormat.Xlsb);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
