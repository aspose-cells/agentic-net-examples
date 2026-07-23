// Title: C# – Validate and Enforce Semantic DocumentVersion in an Aspose.Cells Workbook
// Description: Opens an existing Excel file (or creates a new workbook), reads the Built‑in DocumentVersion property, assigns a default "1.0" when missing, checks the value against a semantic version pattern (major.minor[.patch]) using Regex, throws an error for invalid formats, and saves the workbook. Ensures every exported file carries a correctly formatted version number.
// Keywords: Aspose.Cells | DocumentVersion | semantic version validation | C# | Excel workbook | built‑in document properties | regex version check | default DocumentVersion | .NET | batch Excel processing
// Common Searches: Aspose.Cells validate DocumentVersion C# | set default DocumentVersion in Excel with Aspose | semantic version pattern for DocumentVersion Aspose.Cells | throw error if DocumentVersion is not semver | how to check built‑in document properties in Aspose.Cells
// Developer Intent: Guarantee that a workbook’s DocumentVersion follows a semantic version format before it is saved.
// Use Cases: Prevent downstream applications from failing by ensuring every generated Excel file contains a valid version string. | Automatically apply a fallback version (e.g., 1.0) to legacy workbooks that lack a DocumentVersion property. | Integrate version‑validation into a batch pipeline that creates or modifies large numbers of workbooks with Aspose.Cells.
// AI Prompts: Create a reusable C# method that validates DocumentVersion against a full semver regex (including pre‑release and build metadata) and corrects it when needed. | Show how to log DocumentVersion validation failures with Microsoft.Extensions.Logging while continuing processing of other workbooks. | Provide code that scans a folder of Excel files, validates each DocumentVersion using Aspose.Cells, and moves files with invalid versions to an error directory.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsDocumentVersionValidation
{
    // Opens an existing Excel file (or creates a new workbook), reads the Built‑in DocumentVersion property, assigns a default "1.0" when missing, checks the value against a semantic version pattern (major.minor[.patch]) using Regex, throws an error for invalid formats, and saves the workbook. Ensures every exported file carries a correctly formatted version number.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "ValidatedWorkbook.xlsx";

            try
            {
                Workbook workbook;

                // Ensure the input file exists; if not, create a new workbook
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Optionally set a default DocumentVersion for a new workbook
                    workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";
                }

                // Retrieve the DocumentVersion property (may be null or empty)
                string version = workbook.BuiltInDocumentProperties.DocumentVersion;

                // If version is missing, assign a default value
                if (string.IsNullOrWhiteSpace(version))
                {
                    version = "1.0";
                    workbook.BuiltInDocumentProperties.DocumentVersion = version;
                }

                // Simple semantic version pattern: major.minor[.patch]
                string pattern = @"^\d+\.\d+(\.\d+)?$";

                // Validate the version string
                if (!Regex.IsMatch(version, pattern))
                {
                    throw new InvalidOperationException(
                        $"DocumentVersion \"{version}\" is not a valid semantic version. Expected format: major.minor[.patch]");
                }

                // Save the validated workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\" with valid DocumentVersion \"{version}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
