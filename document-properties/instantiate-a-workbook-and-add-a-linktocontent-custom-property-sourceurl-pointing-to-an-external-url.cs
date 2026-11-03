// Title: Add a URL as a custom document property (SourceUrl) to a new Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add custom document property with external URL | How to set a link-to-content property in an Excel file using Aspose.Cells .NET | Create workbook and store source URL in custom property with Aspose.Cells | Saving a URL in a custom property of an Excel workbook via Aspose.Cells C# | Aspose.Cells add string custom property pointing to website
// Tags: custom property insertion Aspose.Cells | embed external URL in Excel workbook | link-to-content property handling .NET | modify workbook properties Aspose.Cells | store SourceUrl in Excel file

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, adds a custom document property named "SourceUrl" containing an external URL string, ensures the output directory exists, saves the workbook as Output.xlsx, and prints the full file path, with exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a custom document property (as a string; Aspose.Cells does not expose a LinkToContent type directly)
            workbook.CustomDocumentProperties.Add("SourceUrl", "https://www.example.com");

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
