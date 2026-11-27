// Title: Replace deprecated DBConnection.Name values across multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to update DBConnection.Name in all Excel files using Aspose.Cells C# | batch replace old database name in Excel workbook connections with Aspose.Cells | C# program to rename external DB connections in multiple spreadsheets | Aspose.Cells script to change deprecated database name in workbook data connections | process folder of Excel files to modify DBConnection names via Aspose.Cells
// Tags: batch DBConnection name replacement Aspose.Cells | update Excel workbook data connections C# | replace deprecated database name in spreadsheets | process multiple Excel files Aspose.Cells | preserve folder hierarchy when saving modified workbooks

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace ReplaceDeprecatedDbName
{
    // The utility scans a given directory for Excel and compatible spreadsheet files, loads each workbook with Aspose.Cells, iterates through its DataConnections, and replaces any DBConnection.Name containing the deprecated database name with a new name (case‑insensitive). Modified workbooks are saved either in place or to a specified output folder while preserving the original folder structure, and the tool reports the number of files examined and updated.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the folder containing workbooks (required)
        /// args[1] - Deprecated database name to replace (required)
        /// args[2] - New database name (required)
        /// args[3] - Optional: output folder. If omitted, files are overwritten in place.
        /// </param>
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: ReplaceDeprecatedDbName <inputFolder> <oldDbName> <newDbName> [outputFolder]");
                return;
            }

            string inputFolder = args[0];
            string oldDbName = args[1];
            string newDbName = args[2];
            string outputFolder = args.Length >= 4 ? args[3] : null;

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Error: Input folder \"{inputFolder}\" does not exist.");
                return;
            }

            if (outputFolder != null && !Directory.Exists(outputFolder))
            {
                try
                {
                    Directory.CreateDirectory(outputFolder);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: Could not create output folder \"{outputFolder}\": {ex.Message}");
                    return;
                }
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".xlsxml", ".ods", ".csv" };

            string[] files = Directory.GetFiles(inputFolder, "*.*", SearchOption.AllDirectories);
            int processedCount = 0;
            int modifiedCount = 0;

            foreach (string filePath in files)
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // skip non‑Excel files

                try
                {
                    Console.WriteLine($"Processing \"{filePath}\"...");
                    Workbook workbook = new Workbook(filePath);
                    bool workbookModified = false;

                    // Iterate through all external connections
                    ExternalConnectionCollection connections = workbook.DataConnections;
                    for (int i = 0; i < connections.Count; i++)
                    {
                        ExternalConnection conn = connections[i];

                        // We are only interested in DBConnection objects
                        if (conn is DBConnection dbConn)
                        {
                            // The Name property holds the connection name (may be null)
                            string currentName = dbConn.Name ?? string.Empty;

                            // Replace if the old name is found (exact match or substring)
                            if (currentName.Contains(oldDbName, StringComparison.OrdinalIgnoreCase))
                            {
                                string updatedName = currentName.Replace(oldDbName, newDbName, StringComparison.OrdinalIgnoreCase);
                                dbConn.Name = updatedName;
                                workbookModified = true;
                                Console.WriteLine($"  Updated DBConnection name from \"{currentName}\" to \"{updatedName}\" (index {i}).");
                            }
                        }
                    }

                    if (workbookModified)
                    {
                        string savePath;
                        if (outputFolder != null)
                        {
                            // Preserve relative folder structure
                            string relativePath = Path.GetRelativePath(inputFolder, filePath);
                            string targetPath = Path.Combine(outputFolder, relativePath);
                            Directory.CreateDirectory(Path.GetDirectoryName(targetPath));
                            savePath = targetPath;
                        }
                        else
                        {
                            savePath = filePath; // overwrite original
                        }

                        // Save using the same format as the source file
                        workbook.Save(savePath);
                        modifiedCount++;
                        Console.WriteLine($"  Workbook saved to \"{savePath}\".");
                    }

                    processedCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error processing \"{filePath}\": {ex.Message}");
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Processing complete. Files examined: {processedCount}, Workbooks modified: {modifiedCount}.");
        }
    }
}
