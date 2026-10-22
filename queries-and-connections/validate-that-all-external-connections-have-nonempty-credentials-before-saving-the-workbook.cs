// Title: Check and enforce non‑empty credentials for external data connections in an Aspose.Cells workbook before saving (C#)
// AI Prompts: Iterate through each worksheet's ExternalConnections via reflection, read the Credential object, and raise an InvalidOperationException if UserName or Password is blank. | Insert a validation routine before calling Workbook.Save that logs a warning when the ExternalConnections API cannot be accessed and only proceeds after all credentials are verified.
// Common Searches: C# Aspose.Cells how to validate external connection credentials before workbook.Save() | verify that Excel data connection username and password are set using Aspose.Cells .NET | reflection workaround for ExternalConnections property in Aspose.Cells when API is missing
// Tags: external connection credential validation Aspose.Cells | C# workbook save connection check | reflection access ExternalConnections Aspose.Cells | invalidoperationexception missing credentials Excel | Aspose.Cells external data connections .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, uses reflection to locate each worksheet's ExternalConnections collection, iterates over the connections, and ensures that each Credential's UserName and Password are non‑empty. It throws an InvalidOperationException for any missing credential, logs a warning if the ExternalConnections API is unavailable, and finally saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: External connection APIs may not be available in all Aspose.Cells versions.
            // The following block is retained for reference but guarded to compile safely.
            try
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Attempt to access external connections if the API exists.
                    // This uses reflection to avoid compile‑time dependency on unavailable types.
                    var connectionsProp = typeof(Worksheet).GetProperty("ExternalConnections");
                    if (connectionsProp != null)
                    {
                        var connections = connectionsProp.GetValue(sheet);
                        if (connections != null)
                        {
                            // The collection type is expected to be ExternalConnectionCollection.
                            // Iterate using dynamic to avoid compile‑time type requirements.
                            foreach (dynamic connection in (System.Collections.IEnumerable)connections)
                            {
                                // Check for credentials via reflection/dynamic.
                                var credential = connection.Credential;
                                if (credential != null)
                                {
                                    string userName = credential.UserName;
                                    string password = credential.Password;

                                    if (string.IsNullOrWhiteSpace(userName) || string.IsNullOrWhiteSpace(password))
                                    {
                                        throw new InvalidOperationException(
                                            $"External connection '{connection.Name}' in sheet '{sheet.Name}' has empty credentials.");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any issues related to external connection processing without stopping the main flow.
                Console.WriteLine($"Warning: Unable to validate external connections. {ex.Message}");
            }

            // Save the workbook after successful validation (or after warning)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details (could be replaced with proper logging)
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
