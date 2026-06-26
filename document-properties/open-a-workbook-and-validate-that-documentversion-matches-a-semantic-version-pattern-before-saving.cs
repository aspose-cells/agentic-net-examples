using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsDocumentVersionValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook to be opened
            string inputPath = "InputWorkbook.xlsx";

            // Open the workbook using the standard constructor (lifecycle rule)
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the DocumentVersion property
            string version = workbook.BuiltInDocumentProperties.DocumentVersion;

            // Define a simple semantic version pattern: major.minor[.patch]
            string semVerPattern = @"^\d+\.\d+(\.\d+)?$";

            // Validate the version string against the pattern
            if (!Regex.IsMatch(version ?? string.Empty, semVerPattern))
            {
                // If invalid, set a default semantic version
                workbook.BuiltInDocumentProperties.DocumentVersion = "1.0.0";
                Console.WriteLine($"Invalid DocumentVersion '{version}'. Reset to default '1.0.0'.");
            }
            else
            {
                Console.WriteLine($"DocumentVersion '{version}' is valid.");
            }

            // Save the workbook using the standard Save method (lifecycle rule)
            string outputPath = "ValidatedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}