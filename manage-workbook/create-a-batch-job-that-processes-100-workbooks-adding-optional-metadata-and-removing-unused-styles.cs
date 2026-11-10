// Title: C# batch processing of up to 100 Excel workbooks with Aspose.Cells: add custom document properties and clean unused styles
// Common Searches: How to add the same custom document property to many Excel files using Aspose.Cells in C# | Aspose.Cells C# remove unused styles from a workbook to reduce file size | Batch load and save the first 100 Excel workbooks from a folder with Aspose.Cells | Preserve original Excel format while injecting metadata in a C# Aspose.Cells batch job | Create an output folder and process multiple workbooks automatically with Aspose.Cells
// Tags: metadata injection with custom document properties Aspose.Cells C# | remove unused styles to shrink workbook Aspose.Cells | batch processing of Excel workbooks C# Aspose.Cells | limit processing to first 100 files Aspose.Cells | save workbook preserving original format Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookProcessor
{
    // The example loads up to 100 Excel workbooks from a source directory, adds specified custom document properties as metadata, removes any unused styles to minimize file size, and saves each workbook in its original format to an output folder using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source workbooks
            string sourceFolder = @"C:\Workbooks\Input";
            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Workbooks\Output";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Optional metadata to add to each workbook (key-value pairs)
            var optionalMetadata = new Dictionary<string, string>
            {
                { "ProcessedBy", "BatchWorkbookProcessor" },
                { "ProcessedDate", DateTime.UtcNow.ToString("u") }
                // Add more metadata entries as needed
            };

            // Get the first 100 workbook files (any supported format) from the source folder
            string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            int processedCount = 0;

            foreach (string filePath in workbookFiles)
            {
                if (processedCount >= 100)
                    break;

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Add optional metadata as custom document properties
                foreach (var kvp in optionalMetadata)
                {
                    // If the property already exists, update its value; otherwise, add a new one
                    if (workbook.CustomDocumentProperties.Contains(kvp.Key))
                    {
                        workbook.CustomDocumentProperties[kvp.Key].Value = kvp.Value;
                    }
                    else
                    {
                        workbook.CustomDocumentProperties.Add(kvp.Key, kvp.Value);
                    }
                }

                // Remove unused styles to reduce file size
                workbook.RemoveUnusedStyles();

                // Determine output file path (preserve original file name)
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the processed workbook (preserve original format)
                workbook.Save(outputPath);

                processedCount++;
                Console.WriteLine($"Processed {processedCount}: {fileName}");
            }

            Console.WriteLine($"Batch processing completed. Total workbooks processed: {processedCount}");
        }
    }
}
