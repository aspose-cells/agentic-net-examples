using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionValidation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ValidateExternalConnections.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    public class ValidateExternalConnections
    {
        public static void Run()
        {
            try
            {
                // Optional: load a workbook from a template if it exists
                const string templatePath = "Template.xlsx";
                Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

                // Example: add a dummy external connection for demonstration purposes
                // ExternalConnection conn = workbook.DataConnections.Add("MyConnection", "http://example.com/data", "SELECT * FROM Table");
                // conn.CredentialsMethodType = CredentialsMethodType.Integrated;

                bool allConnectionsValid = true;

                // Validate each external connection's credentials
                foreach (ExternalConnection conn in workbook.DataConnections)
                {
                    if (conn.CredentialsMethodType == CredentialsMethodType.None)
                    {
                        Console.WriteLine($"Connection \"{conn.Name}\" has empty credentials.");
                        allConnectionsValid = false;
                    }
                    else
                    {
                        Console.WriteLine($"Connection \"{conn.Name}\" credentials are set to {conn.CredentialsMethodType}.");
                    }
                }

                if (!allConnectionsValid)
                {
                    Console.WriteLine("Workbook contains external connections with empty credentials. Save operation aborted.");
                    return;
                }

                // Save the validated workbook
                const string outputPath = "ValidatedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during validation: {ex.Message}");
            }
        }
    }
}