// Title: Validate Aspose.Cells Workbook DocumentVersion against a semantic version pattern before saving (C#)
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, reads the workbook's DocumentVersion property, checks it with a ^\d+\.\d+\.\d+$ regex, and aborts the save operation if the validation fails. | Refactor the sample to log a custom error message and skip workbook.Save when the DocumentVersion does not conform to a major.minor.patch format.
// Common Searches: how to validate Aspose.Cells workbook DocumentVersion using regex in C# | C# Aspose.Cells check if DocumentVersion follows major.minor.patch before saving | throw exception when Aspose.Cells workbook version is not a semantic version | verify Excel workbook version string matches semantic version pattern with Aspose.Cells
// Tags: Aspose.Cells DocumentVersion regex validation | C# semantic version pattern enforcement | prevent workbook save on version mismatch | Aspose.Cells version property check | InvalidOperationException for invalid DocumentVersion

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook with Aspose.Cells, retrieves the library's DocumentVersion string, validates it against a semantic version regex (major.minor.patch), throws an InvalidOperationException if the pattern does not match, and saves the workbook only when the version is valid.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the Aspose.Cells library version as a string
            string version = typeof(Workbook).Assembly.GetName().Version?.ToString() ?? "0.0.0";

            // Define a semantic version pattern: major.minor.patch (e.g., 1.2.3)
            string semVerPattern = @"^\d+\.\d+\.\d+$";

            // Validate the version against the pattern
            if (!Regex.IsMatch(version, semVerPattern))
                throw new InvalidOperationException($"Workbook version '{version}' does not match the semantic version pattern.");

            // Save the workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
