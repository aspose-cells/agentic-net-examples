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
            // Folder paths – adjust as needed
            string inputFolder = @"C:\Workbooks\Input";
            string outputFolder = @"C:\Workbooks\Output";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process 100 workbooks
            for (int i = 1; i <= 100; i++)
            {
                string inputPath = Path.Combine(inputFolder, $"Workbook{i}.xlsx");
                string outputPath = Path.Combine(outputFolder, $"Workbook{i}_Processed.xlsx");

                // Skip missing input files
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}. Skipping.");
                    continue;
                }

                try
                {
                    // Load workbook
                    using (Workbook wb = new Workbook(inputPath))
                    {
                        // Remove unused styles
                        wb.RemoveUnusedStyles();

                        // Save cleaned workbook
                        wb.Save(outputPath, SaveFormat.Xlsx);
                    }

                    // OPTIONAL: Add custom metadata for even‑indexed workbooks
                    if (i % 2 == 0)
                    {
                        // Prepare metadata options
                        MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
                        // Load metadata for the saved file
                        WorkbookMetadata meta = new WorkbookMetadata(outputPath, metaOptions);

                        // Add a custom document property indicating processing status
                        meta.CustomDocumentProperties.Add("ProcessedByBatch", true);

                        // Save the modified metadata back to the file
                        meta.Save(outputPath);
                    }

                    Console.WriteLine($"Processed workbook {i}: {outputPath}");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing remaining files
                    Console.WriteLine($"Error processing workbook {i}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing of workbooks completed.");
        }
    }
}