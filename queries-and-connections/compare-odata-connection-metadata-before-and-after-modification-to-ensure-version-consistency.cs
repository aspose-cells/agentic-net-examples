// Title: Compare OData WebQueryConnection metadata before and after changes with Aspose.Cells for .NET
// Description: Loads a workbook, extracts the first OData (WebQuery) connection's ConnectionString, Command, IsNew and IsSameSettings, re‑applies the same version token, toggles a non‑version property, saves the file, reloads it, and verifies that the version part of the connection string remains unchanged.
// Keywords: Aspose.Cells | C# OData connection | WebQueryConnection metadata | compare connection string version | verify OData version consistency | external data connections .NET | Aspose.Cells example
// Common Searches: how to read OData connection metadata Aspose.Cells | compare WebQueryConnection properties before and after save | check OData version in connection string using Aspose.Cells | toggle IsSameSettings without changing OData version | Aspose.Cells verify external connection version
// Developer Intent: Validate that modifying a WebQueryConnection does not alter the OData version specified in its connection string.
// Use Cases: Extract and display OData connection string, command, IsNew and IsSameSettings from an existing workbook. | Change a non‑version property of a WebQueryConnection while preserving the original OData version. | Save the workbook, reload it, and confirm that the version token in the connection string is identical to the original.
// AI Prompts: Generate C# code that parses the 'Version' parameter from a WebQueryConnection.ConnectionString and compares it with a reference version. | Refactor the sample to log metadata differences using Aspose.Cells logging utilities. | Create an NUnit test that ensures the OData version remains unchanged after modifying a WebQueryConnection.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Loads a workbook, extracts the first OData (WebQuery) connection's ConnectionString, Command, IsNew and IsSameSettings, re‑applies the same version token, toggles a non‑version property, saves the file, reloads it, and verifies that the version part of the connection string remains unchanged.
    public class ODataConnectionMetadataComparison
    {
        public static void Run()
        {
            try
            {
                const string originalFile = "OriginalWithOData.xlsx";
                const string modifiedFile = "ModifiedWithOData.xlsx";

                // Ensure the original workbook exists
                if (!File.Exists(originalFile))
                {
                    Console.WriteLine($"File not found: {originalFile}");
                    return;
                }

                // Load the original workbook that contains an OData (WebQuery) connection
                Workbook originalWorkbook = new Workbook(originalFile);

                // Find the first WebQueryConnection (used for OData queries)
                WebQueryConnection odataConnection = null;
                foreach (ExternalConnection conn in originalWorkbook.DataConnections)
                {
                    if (conn is WebQueryConnection wqc)
                    {
                        odataConnection = wqc;
                        break;
                    }
                }

                if (odataConnection == null)
                {
                    Console.WriteLine("No OData (WebQuery) connection found in the workbook.");
                    return;
                }

                // Capture metadata before modification
                string beforeConnectionString = odataConnection.ConnectionString;
                string beforeCommand = odataConnection.Command;
                bool beforeIsNew = odataConnection.IsNew;
                bool beforeIsSameSettings = odataConnection.IsSameSettings;

                // Display captured metadata
                Console.WriteLine("=== Metadata BEFORE modification ===");
                Console.WriteLine($"ConnectionString: {beforeConnectionString}");
                Console.WriteLine($"Command: {beforeCommand}");
                Console.WriteLine($"IsNew: {beforeIsNew}");
                Console.WriteLine($"IsSameSettings: {beforeIsSameSettings}");

                // -----------------------------------------------------------------
                // Modify a metadata property while keeping the OData version consistent.
                // Assume the version is specified in the connection string as "Version=4.0".
                // We'll replace the version with the same value to demonstrate a no‑change scenario.
                // -----------------------------------------------------------------
                const string versionKey = "Version=";
                if (!string.IsNullOrEmpty(beforeConnectionString) && beforeConnectionString.Contains(versionKey))
                {
                    // Extract current version substring
                    int startIdx = beforeConnectionString.IndexOf(versionKey) + versionKey.Length;
                    int endIdx = beforeConnectionString.IndexOf(';', startIdx);
                    if (endIdx == -1) endIdx = beforeConnectionString.Length;
                    string currentVersion = beforeConnectionString.Substring(startIdx, endIdx - startIdx);

                    // Re‑apply the same version (no actual change)
                    string newConnectionString = beforeConnectionString.Replace($"{versionKey}{currentVersion}", $"{versionKey}{currentVersion}");
                    odataConnection.ConnectionString = newConnectionString;
                }

                // Optionally toggle a non‑version property to see that version stays the same
                odataConnection.IsSameSettings = !beforeIsSameSettings;

                // Save the modified workbook
                originalWorkbook.Save(modifiedFile);

                // Reload the modified workbook to read back the metadata
                if (!File.Exists(modifiedFile))
                {
                    Console.WriteLine($"Failed to save modified workbook: {modifiedFile}");
                    return;
                }

                Workbook modifiedWorkbook = new Workbook(modifiedFile);
                WebQueryConnection modifiedConnection = null;
                foreach (ExternalConnection conn in modifiedWorkbook.DataConnections)
                {
                    if (conn is WebQueryConnection wqc)
                    {
                        modifiedConnection = wqc;
                        break;
                    }
                }

                if (modifiedConnection == null)
                {
                    Console.WriteLine("No OData (WebQuery) connection found after modification.");
                    return;
                }

                // Capture metadata after modification
                string afterConnectionString = modifiedConnection.ConnectionString;
                string afterCommand = modifiedConnection.Command;
                bool afterIsNew = modifiedConnection.IsNew;
                bool afterIsSameSettings = modifiedConnection.IsSameSettings;

                // Display captured metadata
                Console.WriteLine("\n=== Metadata AFTER modification ===");
                Console.WriteLine($"ConnectionString: {afterConnectionString}");
                Console.WriteLine($"Command: {afterCommand}");
                Console.WriteLine($"IsNew: {afterIsNew}");
                Console.WriteLine($"IsSameSettings: {afterIsSameSettings}");

                // Compare version part of the connection string to ensure consistency
                bool versionConsistent = true;
                if (!string.IsNullOrEmpty(beforeConnectionString) && !string.IsNullOrEmpty(afterConnectionString) &&
                    beforeConnectionString.Contains(versionKey) && afterConnectionString.Contains(versionKey))
                {
                    string GetVersion(string cs)
                    {
                        int s = cs.IndexOf(versionKey) + versionKey.Length;
                        int e = cs.IndexOf(';', s);
                        if (e == -1) e = cs.Length;
                        return cs.Substring(s, e - s);
                    }

                    string beforeVersion = GetVersion(beforeConnectionString);
                    string afterVersion = GetVersion(afterConnectionString);
                    versionConsistent = beforeVersion == afterVersion;
                    Console.WriteLine($"\nVersion consistency check: {(versionConsistent ? "PASS" : "FAIL")}");
                    Console.WriteLine($"Before version: {beforeVersion}, After version: {afterVersion}");
                }
                else
                {
                    Console.WriteLine("\nVersion information not found in connection strings; cannot verify consistency.");
                }
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
