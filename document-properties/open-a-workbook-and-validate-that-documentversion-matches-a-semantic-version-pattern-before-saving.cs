// Title: Validate a custom DocumentVersion property against a semantic version pattern before saving an Excel workbook with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read custom DocumentVersion property and ensure correct format before saving | How to use regex with Aspose.Cells to validate Excel custom property values in .NET | Throw InvalidOperationException when Excel custom document property does not match expected pattern using Aspose.Cells | Validate custom document property version string in an Excel workbook with Aspose.Cells for .NET | Check that DocumentVersion custom property follows major.minor.patch in C# Aspose.Cells code
// Tags: custom property validation Aspose.Cells | regex version string check Excel C# | load workbook verify property before save Aspose.Cells | invalid operation exception for document version Aspose.Cells | semantic version pattern enforcement Excel Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, reads the custom DocumentVersion property if present, validates that its value matches a semantic version pattern (major.minor.patch) using a regular expression, throws an InvalidOperationException on mismatch, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Attempt to read a custom document property named "DocumentVersion"
            string docVersion = string.Empty;
            if (workbook.CustomDocumentProperties.Contains("DocumentVersion"))
            {
                var prop = workbook.CustomDocumentProperties["DocumentVersion"];
                docVersion = prop?.Value?.ToString() ?? string.Empty;
            }

            // Semantic version pattern: major.minor.patch (e.g., 1.2.3)
            const string semVerPattern = @"^\d+\.\d+\.\d+$";

            // Validate the DocumentVersion against the pattern
            if (!Regex.IsMatch(docVersion ?? string.Empty, semVerPattern))
            {
                throw new InvalidOperationException(
                    $"DocumentVersion '{docVersion}' does not match the required semantic version pattern.");
            }

            // Save the workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
