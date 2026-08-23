// Title: Save a modified Excel workbook while preserving the original document properties using Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file with Aspose.Cells, change a cell value, and save it to a new path while keeping the source workbook’s document properties intact. | Use WorkbookMetadata and MetadataOptions to copy the original Excel file’s metadata to a newly saved workbook after making edits in C#.
// Common Searches: how to retain original Excel document properties after editing with Aspose.Cells in C# | Aspose.Cells copy workbook metadata to a new file after modifications | preserve custom properties when saving a modified .xlsx using Aspose.Cells for .NET | C# example for saving workbook with original metadata using Aspose.Cells Metadata API | save modified Excel file without losing built‑in properties Aspose.Cells
// Tags: preserve workbook document properties Aspose.Cells | copy Excel metadata C# Aspose.Cells | save modified workbook with original metadata | WorkbookMetadata usage Aspose.Cells | retain custom document properties .NET Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace PreserveMetadataDemoApp
{
    // The sample loads "original.xlsx", updates cell A1, saves the changes to "modified.xlsx", then reads the source workbook's document properties via WorkbookMetadata and writes them to the new file, ensuring the original metadata is preserved.
    class PreserveMetadataDemo
    {
        public static void Run()
        {
            string sourcePath = "original.xlsx";
            string destPath = "modified.xlsx";

            try
            {
                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the original workbook
                Workbook workbook = new Workbook(sourcePath);

                // Example modification: change the value of cell A1 in the first worksheet
                workbook.Worksheets[0].Cells["A1"].PutValue("Modified content");

                // Save the modified workbook (preserves format)
                workbook.Save(destPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {destPath}");

                // Load original workbook's metadata (document properties)
                MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
                WorkbookMetadata originalMetadata = new WorkbookMetadata(sourcePath, metaOptions);

                // Save metadata to the new file
                originalMetadata.Save(destPath);
                Console.WriteLine("Metadata preserved in the modified workbook.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            PreserveMetadataDemo.Run();
        }
    }
}
