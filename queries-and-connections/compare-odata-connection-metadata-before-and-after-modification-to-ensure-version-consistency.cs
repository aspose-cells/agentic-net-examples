using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsDemo
{
    class ODataMetadataComparison
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "ODataWorkbook.xlsx";
                const string outputPath = "ODataWorkbook_Modified.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load workbook with OData connection
                Workbook workbook = new Workbook(inputPath);

                // Find first OData connection
                ExternalConnection odataConnection = null;
                foreach (ExternalConnection conn in workbook.DataConnections)
                {
                    if (!string.IsNullOrEmpty(conn.ConnectionString) &&
                        conn.ConnectionString.IndexOf("odata", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        odataConnection = conn;
                        break;
                    }
                }

                if (odataConnection == null)
                {
                    Console.WriteLine("No OData connection found in the workbook.");
                    return;
                }

                // Capture original metadata
                string originalConnectionString = odataConnection.ConnectionString;
                string originalCommand = odataConnection.Command;
                bool originalIsNew = odataConnection.IsNew;

                Console.WriteLine("Original Metadata:");
                Console.WriteLine($"ConnectionString: {originalConnectionString}");
                Console.WriteLine($"Command: {originalCommand}");
                Console.WriteLine($"IsNew: {originalIsNew}");

                // Toggle IsNew flag as an example modification
                odataConnection.IsNew = !originalIsNew;

                // Save modified workbook
                workbook.Save(outputPath);

                // Verify saved file exists before reloading
                if (!File.Exists(outputPath))
                {
                    Console.WriteLine($"Failed to create output file '{outputPath}'.");
                    return;
                }

                // Reload and read modified metadata
                Workbook modifiedWorkbook = new Workbook(outputPath);
                ExternalConnection modifiedConnection = null;
                foreach (ExternalConnection conn in modifiedWorkbook.DataConnections)
                {
                    if (!string.IsNullOrEmpty(conn.ConnectionString) &&
                        conn.ConnectionString.IndexOf("odata", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        modifiedConnection = conn;
                        break;
                    }
                }

                if (modifiedConnection == null)
                {
                    Console.WriteLine("Modified OData connection not found.");
                    return;
                }

                // Capture modified metadata
                string modifiedConnectionString = modifiedConnection.ConnectionString;
                string modifiedCommand = modifiedConnection.Command;
                bool modifiedIsNew = modifiedConnection.IsNew;

                Console.WriteLine("\nModified Metadata:");
                Console.WriteLine($"ConnectionString: {modifiedConnectionString}");
                Console.WriteLine($"Command: {modifiedCommand}");
                Console.WriteLine($"IsNew: {modifiedIsNew}");

                // Consistency check (ConnectionString and Command should remain unchanged)
                bool isConsistent = originalConnectionString == modifiedConnectionString &&
                                    originalCommand == modifiedCommand;

                Console.WriteLine($"\nVersion consistency check: {(isConsistent ? "PASS" : "FAIL")}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ODataMetadataComparison.Run();
        }
    }
}