using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class ODataConnectionMetadataComparison
    {
        public static void Run()
        {
            try
            {
                const string originalPath = "OriginalWithOData.xlsx";
                const string modifiedPath = "ModifiedWithOData.xlsx";

                // Verify original workbook exists
                if (!File.Exists(originalPath))
                {
                    Console.WriteLine($"File not found: {originalPath}");
                    return;
                }

                // Load the original workbook that contains the OData connection
                Workbook originalWorkbook = new Workbook(originalPath);

                // Retrieve the first external connection (assumed to be the OData connection)
                if (originalWorkbook.DataConnections.Count == 0)
                {
                    Console.WriteLine("No external connections found in the original workbook.");
                    return;
                }

                ExternalConnection originalConnection = originalWorkbook.DataConnections[0];

                // Capture metadata before modification
                string originalName = originalConnection.Name;
                string originalConnectionString = originalConnection.ConnectionString;
                bool originalIsNew = originalConnection.IsNew;
                bool originalOnlyUseFile = originalConnection.OnlyUseConnectionFile;

                Console.WriteLine("=== Original Connection Metadata ===");
                Console.WriteLine($"Name: {originalName}");
                Console.WriteLine($"ConnectionString: {originalConnectionString}");
                Console.WriteLine($"IsNew: {originalIsNew}");
                Console.WriteLine($"OnlyUseConnectionFile: {originalOnlyUseFile}");
                Console.WriteLine();

                // Modify the connection metadata (e.g., change the OData version in the connection string)
                // Assume the OData version is specified as "Version=1.0" in the connection string
                string modifiedConnectionString = originalConnectionString.Replace("Version=1.0", "Version=2.0");
                originalConnection.ConnectionString = modifiedConnectionString;

                // Optionally toggle other properties to simulate a change
                originalConnection.IsNew = false;
                originalConnection.OnlyUseConnectionFile = !originalOnlyUseFile;

                // Save the workbook after modification
                originalWorkbook.Save(modifiedPath);

                // Verify modified workbook was saved
                if (!File.Exists(modifiedPath))
                {
                    Console.WriteLine($"Failed to create modified file: {modifiedPath}");
                    return;
                }

                // Load the modified workbook
                Workbook modifiedWorkbook = new Workbook(modifiedPath);

                // Retrieve the modified connection
                if (modifiedWorkbook.DataConnections.Count == 0)
                {
                    Console.WriteLine("No external connections found in the modified workbook.");
                    return;
                }

                ExternalConnection modifiedConnection = modifiedWorkbook.DataConnections[0];

                // Capture metadata after modification
                string modifiedName = modifiedConnection.Name;
                string modifiedConnectionStringAfter = modifiedConnection.ConnectionString;
                bool modifiedIsNew = modifiedConnection.IsNew;
                bool modifiedOnlyUseFile = modifiedConnection.OnlyUseConnectionFile;

                Console.WriteLine("=== Modified Connection Metadata ===");
                Console.WriteLine($"Name: {modifiedName}");
                Console.WriteLine($"ConnectionString: {modifiedConnectionStringAfter}");
                Console.WriteLine($"IsNew: {modifiedIsNew}");
                Console.WriteLine($"OnlyUseConnectionFile: {modifiedOnlyUseFile}");
                Console.WriteLine();

                // Helper to extract version from a connection string
                string GetVersion(string connStr)
                {
                    const string token = "Version=";
                    int idx = connStr.IndexOf(token, StringComparison.OrdinalIgnoreCase);
                    if (idx < 0) return "NotSpecified";
                    int start = idx + token.Length;
                    int end = connStr.IndexOf(';', start);
                    if (end < 0) end = connStr.Length;
                    return connStr.Substring(start, end - start).Trim();
                }

                string originalVersion = GetVersion(originalConnectionString);
                string modifiedVersion = GetVersion(modifiedConnectionStringAfter);

                Console.WriteLine("=== Version Consistency Check ===");
                Console.WriteLine($"Original Version: {originalVersion}");
                Console.WriteLine($"Modified Version: {modifiedVersion}");
                Console.WriteLine(originalVersion == modifiedVersion
                    ? "Version is consistent."
                    : "Version mismatch detected.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ODataConnectionMetadataComparison.Run();
        }
    }
}