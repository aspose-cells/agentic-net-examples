// Title: C# example: Validate external data connections have non‑empty credentials before saving an Excel workbook with Aspose.Cells
// Common Searches: how to check external connection credentials in Aspose.Cells before saving workbook | C# validate ODBC/OLEDB connection strings for empty User ID or Password using Aspose.Cells | Aspose.Cells .NET ensure DataConnections contain required parameters before workbook.save | verify external data connections credentials in Excel file with Aspose.Cells C# | prevent saving workbook when external connection credentials are missing in Aspose.Cells
// Tags: external connection credential validation Aspose.Cells | DataConnections connection string check C# | non‑empty User ID Password verification .NET | Aspose.Cells workbook save guard | ODBC OLEDB connection integrity check

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace ValidateExternalConnectionsDemo
{
    // The sample loads an input .xlsx file with Aspose.Cells, iterates through its DataConnections collection, and verifies that each connection string contains non‑empty User ID, UserName, Password, or Pwd parameters. If any required credential is missing or empty, the validation fails and the workbook is not saved; otherwise the workbook is saved to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: ValidateExternalConnectionsDemo <input.xlsx> <output.xlsx>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Validate external connections
                if (!ValidateExternalConnections(workbook))
                {
                    Console.WriteLine("Validation failed: one or more external connections have empty credentials.");
                    return;
                }

                // Save the workbook (validation passed)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static bool ValidateExternalConnections(Workbook workbook)
        {
            // DataConnections collection holds all external connections (ODBC, OLEDB, Web, etc.)
            ExternalConnectionCollection connections = workbook.DataConnections;

            if (connections == null || connections.Count == 0)
            {
                Console.WriteLine("No external connections found in the workbook.");
                return true; // Nothing to validate
            }

            bool allValid = true;
            var requiredParams = new[] { "User ID", "UserName", "Password", "Pwd" };

            for (int i = 0; i < connections.Count; i++)
            {
                ExternalConnection conn = connections[i];
                // Use SourceType instead of the obsolete Type property
                Console.WriteLine($"Validating connection #{i + 1}: Name = '{conn.Name}', SourceType = {conn.SourceType}");

                // Retrieve the connection string; it may be empty for some connection types
                string connString = conn.ConnectionString ?? string.Empty;
                var missingOrEmpty = new List<string>();

                foreach (string paramName in requiredParams)
                {
                    // Simple check: look for "paramName=" pattern and ensure a value follows
                    string search = paramName + "=";
                    int idx = connString.IndexOf(search, StringComparison.OrdinalIgnoreCase);
                    if (idx == -1)
                    {
                        missingOrEmpty.Add(paramName);
                    }
                    else
                    {
                        int valueStart = idx + search.Length;
                        int valueEnd = connString.IndexOf(';', valueStart);
                        string value = valueEnd == -1
                            ? connString.Substring(valueStart).Trim()
                            : connString.Substring(valueStart, valueEnd - valueStart).Trim();

                        if (string.IsNullOrWhiteSpace(value))
                            missingOrEmpty.Add(paramName);
                    }
                }

                if (missingOrEmpty.Count > 0)
                {
                    allValid = false;
                    Console.WriteLine($"  Invalid: Missing or empty credential parameters: {string.Join(", ", missingOrEmpty)}");
                }
                else
                {
                    Console.WriteLine("  Valid: All required credential parameters are present.");
                }
            }

            return allValid;
        }
    }
}
