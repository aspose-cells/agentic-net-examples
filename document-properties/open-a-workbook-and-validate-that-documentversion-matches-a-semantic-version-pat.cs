using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsDocumentVersionValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "ValidatedWorkbook.xlsx";

            // Open the workbook (uses the provided Workbook(string) constructor)
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the DocumentVersion built‑in property
            string version = workbook.BuiltInDocumentProperties.DocumentVersion;

            // Define a simple semantic version pattern: major.minor.patch (e.g., 1.0.0)
            string semVerPattern = @"^\d+\.\d+\.\d+$";

            // Validate the version string against the pattern
            if (!Regex.IsMatch(version ?? string.Empty, semVerPattern))
            {
                // If the version is missing or does not match, set a default valid version
                workbook.BuiltInDocumentProperties.DocumentVersion = "1.0.0";
                Console.WriteLine($"DocumentVersion \"{version}\" is invalid. Reset to default \"1.0.0\".");
            }
            else
            {
                Console.WriteLine($"DocumentVersion \"{version}\" is valid.");
            }

            // Save the workbook (uses the provided Workbook.Save(string) method)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
    }
}