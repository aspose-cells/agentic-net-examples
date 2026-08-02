using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DocumentVersionValidator
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the existing workbook
            string inputPath = "InputWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Open the workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the DocumentVersion property
            string version = workbook.BuiltInDocumentProperties.DocumentVersion;

            // Define a semantic version pattern: major.minor[.patch] (e.g., 1.0 or 2.5.3)
            string pattern = @"^\d+\.\d+(\.\d+)?$";

            // Validate the version string against the pattern
            if (!Regex.IsMatch(version ?? string.Empty, pattern))
            {
                throw new InvalidOperationException(
                    $"DocumentVersion \"{version}\" does not match the required semantic version pattern.");
            }

            // If validation passes, save the workbook (save rule)
            string outputPath = "ValidatedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\" with valid DocumentVersion \"{version}\".");
        }
    }
}