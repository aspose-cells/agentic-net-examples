// Title: Rename the first DataConnection in an XLSB workbook using Aspose.Cells for .NET and log the modification timestamp
// Common Searches: C# Aspose.Cells change the name of a DataConnection in an XLSB file | How to log the time when a workbook data connection is renamed with Aspose.Cells | Load an XLSB workbook, modify its first data connection, and save it using Aspose.Cells | Aspose.Cells .NET check for existing DataConnections before renaming | Update DBConnection name in XLSB and capture ISO‑8601 timestamp in C#
// Tags: data connection renaming Aspose.Cells | XLSB workbook data connection update | timestamp logging for data connection changes C# | load and save XLSB with Aspose.Cells | validate presence of data connections Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSB workbook, ensures it contains worksheets, checks for workbook‑level data connections, renames the first DataConnection to a new name, writes an ISO‑8601 timestamp to the console, and saves the modified workbook back to XLSB, with comprehensive error handling for loading, modification, and saving steps.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsb";
        const string outputPath = "output.xlsb";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;

        try
        {
            // Load the XLSB workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Ensure there is at least one worksheet
        if (workbook.Worksheets.Count == 0)
        {
            Console.WriteLine("The workbook contains no worksheets.");
            return;
        }

        // Check for data connections at the workbook level
        if (workbook.DataConnections != null && workbook.DataConnections.Count > 0)
        {
            try
            {
                // Change the name of the first data connection
                var dataConnection = workbook.DataConnections[0];
                dataConnection.Name = "NewConnectionName";

                // Log the modification timestamp
                Console.WriteLine($"DataConnection name changed at {DateTime.Now:O}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to modify data connection: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("No data connections found in the workbook.");
        }

        try
        {
            // Save the modified workbook back to XLSB format
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
