// Title: Validate a custom "Version" document property as a semantic version before saving an Excel workbook with Aspose.Cells for .NET
// Common Searches: how to verify that an Excel custom document property follows semantic versioning using Aspose.Cells in C# | Aspose.Cells C# check custom property format before calling Workbook.Save | C# regex to enforce semver pattern on a custom property in an Excel file with Aspose.Cells | load workbook, add default Version property, ensure it matches semver and then save using Aspose.Cells | prevent saving an Excel workbook when the custom Version property is not a valid semantic version in C#
// Tags: custom document property semver validation Aspose.Cells | conditional workbook save based on property format C# | regex semantic version check Excel property Aspose.Cells | fallback workbook creation when input file missing or password protected | default version property initialization Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an existing Excel file if it exists, creates a new workbook when the file is missing or password‑protected, ensures a "Version" custom document property is present (adding a default "1.0.0" if needed), validates the property against a semantic version regular expression, and saves the workbook only when the validation succeeds.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if the file is present
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex) when (ex.Message.Contains("password"))
                {
                    // If the file is password‑protected, create a new workbook instead
                    workbook = new Workbook();
                }
            }
            else
            {
                // Create a new workbook when the input file does not exist
                workbook = new Workbook();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Retrieve or create the "Version" custom document property
        string version;
        if (workbook.CustomDocumentProperties.Contains("Version"))
        {
            version = workbook.CustomDocumentProperties["Version"]?.ToString() ?? string.Empty;
        }
        else
        {
            version = "1.0.0";
            workbook.CustomDocumentProperties.Add("Version", version);
        }

        // Semantic version pattern (e.g., 1.2.3, 1.2.3-alpha, 1.2.3+build)
        const string semVerPattern = @"^\d+\.\d+\.\d+(-[0-9A-Za-z-.]+)?(\+[0-9A-Za-z-.]+)?$";

        // Validate the version string
        if (!Regex.IsMatch(version, semVerPattern))
        {
            Console.WriteLine($"DocumentVersion '{version}' does not match the required semantic version format.");
            return;
        }

        // Save the workbook after successful validation
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
