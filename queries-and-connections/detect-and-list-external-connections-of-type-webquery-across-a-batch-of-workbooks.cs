// Title: Batch detect and list WebQuery external connections in Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells find WebQuery data connections in multiple Excel files | how to list WebQueryConnection details for all workbooks in a folder using Aspose.Cells | enumerate external WebQuery connections across a directory of .xlsx files with Aspose.Cells .NET | retrieve URL and settings of WebQuery connections from Excel workbooks programmatically | scan Excel workbooks for WebQuery external connections using Aspose.Cells API
// Tags: batch enumerate WebQuery connections Aspose.Cells | extract WebQueryConnection properties .NET | scan Excel files for external data connections C# | list DataConnections in .xlsx workbooks using Aspose.Cells | detect WebQuery external connections programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The C# program scans a given folder (or a default "Workbooks" directory) for Excel files, loads each workbook with Aspose.Cells, iterates its DataConnections collection, and outputs details of any connections whose ClassType is WebQuery. It prints connection name, type, command, refresh settings, connection string, and WebQuery‑specific properties such as Url, IsHtmlTables, IsXml, HtmlFormat, and IsConsecutive, while handling missing files and load errors.
class Program
{
    static void Main(string[] args)
    {
        // Determine the folder that contains the workbooks.
        // If a folder path is supplied as the first argument, use it; otherwise use a default folder named "Workbooks" in the current directory.
        string folderPath = args.Length > 0 ? args[0] : Path.Combine(Directory.GetCurrentDirectory(), "Workbooks");

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Error: The folder \"{folderPath}\" does not exist.");
            return;
        }

        // Get all Excel files in the folder (including subfolders).
        string[] workbookFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
        if (workbookFiles.Length == 0)
        {
            Console.WriteLine($"No workbook files found in \"{folderPath}\".");
            return;
        }

        Console.WriteLine($"Scanning {workbookFiles.Length} file(s) for WebQuery external connections...\n");

        foreach (string filePath in workbookFiles)
        {
            // Only attempt to load files with known Excel extensions.
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                continue;

            Console.WriteLine($"Processing workbook: {Path.GetFileName(filePath)}");

            try
            {
                // Ensure the file actually exists before loading.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"  File not found: {filePath}");
                    continue;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Access the collection of external connections.
                ExternalConnectionCollection connections = workbook.DataConnections;

                bool anyWebQuery = false;
                for (int i = 0; i < connections.Count; i++)
                {
                    ExternalConnection conn = connections[i];

                    // Identify WebQuery connections by their ClassType.
                    if (conn.ClassType == ExternalConnectionClassType.WebQuery)
                    {
                        anyWebQuery = true;
                        Console.WriteLine($"  Connection #{i + 1}:");
                        Console.WriteLine($"    Name               : {conn.Name}");
                        Console.WriteLine($"    ClassType          : {conn.ClassType}");
                        Console.WriteLine($"    SourceType         : {conn.SourceType}");
                        Console.WriteLine($"    Command            : {conn.Command}");
                        Console.WriteLine($"    BackgroundRefresh  : {conn.BackgroundRefresh}");
                        Console.WriteLine($"    ConnectionString   : {conn.ConnectionString}");

                        // Attempt to cast to WebQueryConnection for WebQuery‑specific properties.
                        if (conn is WebQueryConnection webConn)
                        {
                            Console.WriteLine($"    URL                : {webConn.Url}");
                            Console.WriteLine($"    IsHtmlTables       : {webConn.IsHtmlTables}");
                            Console.WriteLine($"    IsXml              : {webConn.IsXml}");
                            Console.WriteLine($"    HtmlFormat         : {webConn.HtmlFormat}");
                            Console.WriteLine($"    IsConsecutive      : {webConn.IsConsecutive}");
                        }
                    }
                }

                if (!anyWebQuery)
                {
                    Console.WriteLine("  No WebQuery external connections found in this workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Failed to process \"{Path.GetFileName(filePath)}\": {ex.Message}");
            }

            Console.WriteLine(); // Blank line between workbooks.
        }

        Console.WriteLine("Scanning completed.");
    }
}
