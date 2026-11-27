// Title: Export all external connection names from an Excel workbook to a plain‑text file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells list external data connections in an existing workbook | How to write Excel workbook data connection names to a .txt file with Aspose.Cells | Retrieve external connection names from an .xlsx file using Aspose.Cells API | Save workbook.DataConnections names to plain text in a .NET console app | Export external connection list from Excel using Aspose.Cells for .NET
// Tags: Aspose.Cells export external connection names | C# write DataConnections to text file | Workbook.DataConnections list extraction | Excel external connections to plain text | Aspose.Cells external connection enumeration

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// A console utility that takes an input Excel workbook path and an output text file path, loads the workbook with Aspose.Cells, iterates over its DataConnections collection, extracts each connection's Name, and writes the names line‑by‑line to the specified plain‑text file while handling argument validation and file errors.
class Program
{
    static void Main(string[] args)
    {
        // Expect exactly two arguments: input workbook path and output text file path
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: Program.exe <inputWorkbookPath> <outputTextFilePath>");
            return;
        }

        string inputWorkbookPath = args[0];
        string outputTextFilePath = args[1];

        // Validate input workbook existence
        if (!File.Exists(inputWorkbookPath))
        {
            Console.WriteLine($"Error: The workbook file \"{inputWorkbookPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputWorkbookPath);

            // Retrieve the collection of external connections
            ExternalConnectionCollection externalConnections = workbook.DataConnections;

            // Prepare a list to hold connection names
            List<string> connectionNames = new List<string>();

            // Iterate through each external connection and collect its name
            foreach (ExternalConnection connection in externalConnections)
            {
                // The Name property holds the connection's display name.
                // Guard against null values just in case.
                string name = connection.Name ?? string.Empty;
                connectionNames.Add(name);
            }

            // Write the names to the specified plain‑text file (one name per line)
            File.WriteAllLines(outputTextFilePath, connectionNames);

            Console.WriteLine($"Successfully exported {connectionNames.Count} external connection name(s) to \"{outputTextFilePath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
