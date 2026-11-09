// Title: Batch replace external workbook link paths in multiple Excel .xlsx files with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch update external link paths for a folder of .xlsx files | How to change ExternalFileName for all external links in multiple workbooks using Aspose.Cells | Programmatically replace old external workbook references with new locations in Excel files using Aspose.Cells .NET | Iterate through Worksheets.ExternalLinks collection to modify link paths in bulk with C# | Update external data source paths in a batch of Excel workbooks using Aspose.Cells API
// Tags: batch external link path update Aspose.Cells | modify Worksheets.ExternalLinks C# | replace ExternalFileName in .xlsx files | process multiple Excel workbooks Aspose.Cells | update linked workbook references programmatically

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified input folder for .xlsx workbooks, loads each with Aspose.Cells, iterates the Worksheets.ExternalLinks collection, replaces any ExternalFileName that starts with an old base path with a new base path, and saves the modified workbook to an output directory.
class BatchExternalLinkUpdater
{
    static void Main()
    {
        // Folder containing the source workbooks
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the updated workbooks will be saved
        string outputFolder = @"C:\UpdatedWorkbooks";

        // Define old and new base paths for external links
        // Example: replace "C:\OldData\" with "D:\NewData\"
        string oldBasePath = @"C:\OldData\";
        string newBasePath = @"D:\NewData\";

        try
        {
            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string filePath in workbookFiles)
            {
                // Verify the source file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through all external links in the workbook
                    foreach (var linkObj in workbook.Worksheets.ExternalLinks)
                    {
                        // Use dynamic to access the ExternalFileName property (avoids compile‑time binding issues)
                        dynamic link = linkObj;

                        // Get the current source file path of the external link
                        string currentPath = link.ExternalFileName as string;

                        // If the path starts with the old base path, replace it with the new base path
                        if (!string.IsNullOrEmpty(currentPath) &&
                            currentPath.StartsWith(oldBasePath, StringComparison.OrdinalIgnoreCase))
                        {
                            string updatedPath = newBasePath + currentPath.Substring(oldBasePath.Length);
                            link.ExternalFileName = updatedPath;
                        }
                    }

                    // Save the updated workbook to the output folder, preserving the original file name
                    string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputFilePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
