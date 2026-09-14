// Title: Add a link-to-content custom document property named SourceUrl to a new Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate a new Workbook, insert a string‑type custom document property called SourceUrl with an external URL, and save the file as Output.xlsx using Aspose.Cells for .NET. | Using Aspose.Cells, create an empty Excel file, add a link-to-content custom property that points to a web resource, and ensure the output directory exists before saving. | Write C# code that adds a custom document property storing a source URL to a workbook's metadata and persists the workbook to disk with Aspose.Cells.
// Common Searches: how to add a custom document property with a URL in Aspose.Cells C# | Aspose.Cells set link-to-content property in new workbook | store external source link in Excel file metadata using .NET | C# create workbook and add string custom property Aspose.Cells example
// Tags: add custom document property Aspose.Cells | link-to-content custom property Excel .NET | store external URL in workbook metadata | Aspose.Cells save workbook to .xlsx

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new empty Workbook, adds a custom document property named "SourceUrl" containing an external URL string, ensures the output directory exists, saves the workbook as Output.xlsx, and handles any exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Add a custom document property (string type) containing the source URL
                // The overload without specifying the type defaults to string, which is safe across all Aspose.Cells versions
                workbook.CustomDocumentProperties.Add("SourceUrl", "https://www.example.com/data");

                // Define output file path
                string outputPath = "Output.xlsx";

                // Ensure the directory for the output file exists (if any)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
