// Title: Securely Validate External Connection Paths in an Aspose.Cells Workbook Using a Whitelisted Base Folder (C#)
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, iterates over workbook.DataConnections, resolves each connection.SourceFile to an absolute path, and retains only those paths that reside inside a predefined safe directory. | Write a C# helper method that normalizes a file path, checks it against a whitelist folder, clears the ExternalConnection.SourceFile when the check fails, and logs the action for each connection. | Show how to safely save the sanitized workbook to a designated output folder after processing external connections, including error handling for missing files and invalid paths.
// Common Searches: how to restrict Aspose.Cells external data connections to a specific directory in C# | C# code to prevent directory traversal in Excel workbook external links using Aspose.Cells | validate source file paths of DataConnections in Aspose.Cells before saving workbook | Aspose.Cells secure handling of external connection file paths in .NET | example of whitelisting external Excel file locations with Aspose.Cells C#
// Tags: Aspose.Cells external connection path validation | C# whitelist directory for Excel data connections | prevent directory traversal Aspose.Cells | secure external file handling in Excel workbook | sanitize ExternalConnection.SourceFile C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The example defines a trusted base directory, loads an Excel workbook with Aspose.Cells, iterates through its DataConnections, resolves each connection's source file to an absolute path, verifies the path stays within the allowed folder to block directory‑traversal or unauthorized access, clears unsafe paths, and finally saves the sanitized workbook to a secure location.
class SecureExternalPathHandler
{
    static void Main()
    {
        // Define a base directory that is considered safe for external files.
        string baseDirectory = Path.GetFullPath(@"C:\AllowedExternalFiles");

        // Path to the workbook that will be processed.
        string inputPath = Path.Combine(baseDirectory, "input.xlsx");

        // Verify that the input workbook exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input workbook not found: {inputPath}");
            return;
        }

        // Load the workbook using Aspose.Cells (create/load rule).
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all external connections in the workbook.
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            string sourceFile = connection.SourceFile;
            if (string.IsNullOrEmpty(sourceFile))
                continue; // No external file associated with this connection.

            // Resolve the provided path to an absolute path.
            string absolutePath;
            try
            {
                absolutePath = Path.GetFullPath(sourceFile);
            }
            catch (Exception ex)
            {
                // If the path cannot be resolved, skip this connection.
                Console.WriteLine($"Invalid path '{sourceFile}': {ex.Message}");
                continue;
            }

            // Ensure the absolute path is within the allowed base directory.
            // This prevents directory traversal attacks and unauthorized file access.
            if (!absolutePath.StartsWith(baseDirectory, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Blocked external path outside allowed directory: {absolutePath}");
                // Optionally clear the unsafe path to avoid later failures.
                connection.SourceFile = string.Empty;
                continue;
            }

            // Normalize the path (optional) and assign it back to the connection.
            connection.SourceFile = absolutePath;
        }

        // Define the output path for the modified workbook.
        string outputPath = Path.Combine(baseDirectory, "output.xlsx");

        // Save the workbook using Aspose.Cells (save rule).
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved securely to: {outputPath}");
    }
}
