using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionValidation
{
    public class ValidateExternalConnections
    {
        public static void Run(string inputPath, string outputPath)
        {
            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Validate each external connection's credentials
                foreach (ExternalConnection connection in workbook.DataConnections)
                {
                    // Use the modern CredentialsMethodType property
                    if (connection.CredentialsMethodType == CredentialsMethodType.None)
                    {
                        throw new InvalidOperationException(
                            $"External connection '{connection.Name}' has empty credentials.");
                    }
                }

                // Save the workbook after successful validation
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Log the error to the console
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Default file paths; can be overridden via command‑line arguments
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            if (args.Length > 0) inputPath = args[0];
            if (args.Length > 1) outputPath = args[1];

            ValidateExternalConnections.Run(inputPath, outputPath);
        }
    }
}