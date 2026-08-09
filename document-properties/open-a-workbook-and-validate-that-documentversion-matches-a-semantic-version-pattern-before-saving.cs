// Title: C# – Validate and Fix Workbook DocumentVersion with Aspose.Cells
// Description: Opens an Excel workbook using Aspose.Cells for .NET, reads the Built‑in DocumentVersion property, checks it against a semantic version regex (major.minor[.patch]), substitutes an invalid value with the default "1.0.0", and saves the updated file.
// Keywords: Aspose.Cells | C# | DocumentVersion | semantic version | regex validation | Excel metadata | built‑in document properties | workbook save | version compliance
// Common Searches: Aspose.Cells validate DocumentVersion C# | check Excel file version format with Aspose | regex semantic version for Excel workbook | set default DocumentVersion Aspose.Cells | enforce version pattern in Excel metadata
// Developer Intent: Ensure a workbook's DocumentVersion follows a semantic version pattern before saving.
// Use Cases: Validate the DocumentVersion of a single workbook and correct it automatically. | Integrate version checking into an automated reporting pipeline to keep metadata consistent across generated Excel files. | Process a batch of .xlsx files, fix non‑conforming DocumentVersion values, and log the changes for audit purposes.
// AI Prompts: Create a C# method that takes a Workbook object, validates its DocumentVersion against a semantic version regex, and sets it to "1.0.0" when invalid. | Write a script that scans a folder of Excel files, uses Aspose.Cells to enforce a proper DocumentVersion format, updates any mismatches, and generates a summary report. | Generate a reusable Aspose.Cells utility class that encapsulates DocumentVersion validation, defaulting logic, and optional logging.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsVersionValidation
{
    // Opens an Excel workbook using Aspose.Cells for .NET, reads the Built‑in DocumentVersion property, checks it against a semantic version regex (major.minor[.patch]), substitutes an invalid value with the default "1.0.0", and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook to be opened
            string inputPath = "InputWorkbook.xlsx";

            // Open the workbook using the provided constructor rule
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the DocumentVersion property
            string version = workbook.BuiltInDocumentProperties.DocumentVersion;

            // Define a simple semantic version pattern: major.minor[.patch]
            string pattern = @"^\d+\.\d+(\.\d+)?$";

            // Validate the version string against the pattern
            if (!Regex.IsMatch(version ?? string.Empty, pattern))
            {
                // If the version does not match, set a default valid version
                workbook.BuiltInDocumentProperties.DocumentVersion = "1.0.0";
                Console.WriteLine($"Invalid DocumentVersion '{version}'. Reset to default '1.0.0'.");
            }
            else
            {
                Console.WriteLine($"DocumentVersion '{version}' is valid.");
            }

            // Save the workbook using the provided Save method
            string outputPath = "ValidatedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
