using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

class DocumentVersionValidator
{
    static void Main()
    {
        // Load an existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Retrieve the DocumentVersion property
        string version = workbook.BuiltInDocumentProperties.DocumentVersion;

        // Semantic version pattern: major.minor.patch with optional prerelease and build metadata
        string semVerPattern = @"^\d+\.\d+\.\d+(-[0-9A-Za-z-.]+)?(\+[0-9A-Za-z-.]+)?$";

        // Validate the version string
        if (!Regex.IsMatch(version ?? string.Empty, semVerPattern))
        {
            Console.WriteLine($"Invalid DocumentVersion '{version}'. Setting default version.");
            workbook.BuiltInDocumentProperties.DocumentVersion = "1.0.0";
        }

        // Save the workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}