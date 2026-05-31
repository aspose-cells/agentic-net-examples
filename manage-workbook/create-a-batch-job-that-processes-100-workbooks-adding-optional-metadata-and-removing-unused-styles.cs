using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace BatchWorkbookProcessor
{
    class Program
    {
        static void Main()
        {
            // Directory containing the source workbooks
            string sourceDir = @"C:\Workbooks\Source";
            // Directory where processed workbooks will be saved
            string outputDir = @"C:\Workbooks\Processed";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDir);

            // Process 100 workbooks named Workbook1.xlsx ... Workbook100.xlsx
            for (int i = 1; i <= 100; i++)
            {
                string inputPath = Path.Combine(sourceDir, $"Workbook{i}.xlsx");
                string outputPath = Path.Combine(outputDir, $"Workbook{i}_Processed.xlsx");
                string tempMetaPath = Path.Combine(outputDir, $"Workbook{i}_TempMeta.xlsx");

                try
                {
                    // Verify source file exists
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"Source file not found: {inputPath}");
                        continue;
                    }

                    // Load the workbook
                    using (Workbook wb = new Workbook(inputPath))
                    {
                        // OPTIONAL: Add custom metadata using WorkbookMetadata
                        MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
                        WorkbookMetadata metadata = new WorkbookMetadata(inputPath, metaOptions);

                        // Add custom document properties
                        metadata.CustomDocumentProperties.Add("ProcessedOn", DateTime.UtcNow.ToString("o"));
                        metadata.CustomDocumentProperties.Add("ProcessedBy", "BatchJob");

                        // Save metadata to a temporary file
                        metadata.Save(tempMetaPath);

                        // Reload workbook with updated metadata
                        wb.Dispose(); // Dispose original workbook before reloading
                        using (Workbook wbMeta = new Workbook(tempMetaPath))
                        {
                            // Remove all unused styles
                            wbMeta.RemoveUnusedStyles();

                            // Save the final processed workbook
                            wbMeta.Save(outputPath);
                        }
                    }

                    // Clean up temporary file
                    if (File.Exists(tempMetaPath))
                    {
                        File.Delete(tempMetaPath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing Workbook{i}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing of 100 workbooks completed.");
        }
    }
}