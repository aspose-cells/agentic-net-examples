// Title: Batch rename DBConnection.Name in all XLS workbooks within a folder using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells batch rename DBConnection objects in multiple XLS files | How to change external data connection names for Excel 97-2003 workbooks with Aspose.Cells | Save modified XLS workbook in place after updating DBConnection using Aspose.Cells .NET | Iterate workbook.DataConnections collection to update DBConnection.Name in C#
// Tags: bulk update DBConnection names with Aspose.Cells | modify external data connections in XLS using Aspose.Cells | overwrite original XLS file with Aspose.Cells | loop through workbook.DataConnections in C# | handle multiple Excel 97-2003 files Aspose.Cells | programmatic DBConnection.Name assignment

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The example scans a directory for *.xls files, loads each workbook with Aspose.Cells, walks through its DataConnections collection, assigns a new Name to any DBConnection based on the file name, and overwrites the original XLS file using XlsSaveOptions while reporting progress and handling errors.
class Program
{
    static void Main(string[] args)
    {
        // Determine the directory containing the XLS files.
        // If a path is supplied as a command‑line argument use it, otherwise use the current directory.
        string directoryPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

        if (!Directory.Exists(directoryPath))
        {
            Console.WriteLine($"Error: The directory \"{directoryPath}\" does not exist.");
            return;
        }

        // Get all *.xls files (Excel 97‑2003 format) in the directory (non‑recursive).
        string[] xlsFiles = Directory.GetFiles(directoryPath, "*.xls", SearchOption.TopDirectoryOnly);

        if (xlsFiles.Length == 0)
        {
            Console.WriteLine($"No XLS files found in \"{directoryPath}\".");
            return;
        }

        Console.WriteLine($"Found {xlsFiles.Length} XLS file(s) in \"{directoryPath}\".");

        // Prepare save options for XLS format.
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        foreach (string filePath in xlsFiles)
        {
            try
            {
                Console.WriteLine($"\nProcessing file: {Path.GetFileName(filePath)}");

                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Access the collection of external data connections.
                ExternalConnectionCollection connections = workbook.DataConnections;

                bool anyDbConnectionUpdated = false;

                // Iterate over all connections.
                for (int i = 0; i < connections.Count; i++)
                {
                    ExternalConnection conn = connections[i];

                    // We are interested only in DBConnection objects.
                    if (conn is DBConnection dbConn)
                    {
                        // Build a new name – you can adjust the naming scheme as required.
                        string newName = $"{Path.GetFileNameWithoutExtension(filePath)}_Conn";

                        // Update the Name property.
                        // ExternalConnection (base class) exposes a Name property.
                        dbConn.Name = newName;

                        Console.WriteLine($"  Updated DBConnection at index {i} to Name = \"{newName}\"");
                        anyDbConnectionUpdated = true;
                    }
                }

                if (anyDbConnectionUpdated)
                {
                    // Save the workbook back to the same file (in‑place update).
                    workbook.Save(filePath, saveOptions);
                    Console.WriteLine("  Changes saved successfully.");
                }
                else
                {
                    Console.WriteLine("  No DBConnection objects found – file left unchanged.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error processing \"{Path.GetFileName(filePath)}\": {ex.Message}");
            }
        }

        Console.WriteLine("\nProcessing completed.");
    }
}
