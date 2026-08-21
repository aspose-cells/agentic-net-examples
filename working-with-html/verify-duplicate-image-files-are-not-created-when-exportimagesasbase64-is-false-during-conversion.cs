// Title: Detect duplicate image files when exporting Excel to HTML with ExportImagesAsBase64 = false (Aspose.Cells .NET)
// Description: The sample creates a temporary PNG, inserts it into a new workbook, saves the workbook as HTML with images written to a separate folder, then scans that folder, groups file names case‑insensitively, and reports any naming collisions.
// Keywords: Aspose.Cells | .NET | C# | HTML export | ExportImagesAsBase64 | duplicate image detection | image folder | Excel to HTML conversion | picture handling | file name collision
// Common Searches: Aspose.Cells HTML export duplicate images | C# check for duplicate image files after saving workbook as HTML | ExportImagesAsBase64 false image folder Aspose.Cells | how to prevent image name collisions in Aspose.Cells HTML output | verify image export integrity Aspose.Cells .NET
// Developer Intent: Ensure that saving a workbook to HTML with ExportImagesAsBase64 disabled does not produce repeated image files.
// Use Cases: Automated validation of HTML conversion pipelines for Excel reports. | Detecting naming conflicts when multiple pictures are embedded before export. | Maintaining clean asset directories for web deployment of converted workbooks.
// AI Prompts: Generate a unit test in C# that fails if any duplicate image files appear after an Aspose.Cells HTML export with ExportImagesAsBase64 set to false. | Rewrite the duplicate‑check logic using a HashSet for O(n) performance and provide the updated code snippet. | Explain the effect of setting HtmlSaveOptions.ExportImagesAsBase64 to true on the output structure and why duplicate file concerns disappear.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsImageDuplicationCheck
{
    // The sample creates a temporary PNG, inserts it into a new workbook, saves the workbook as HTML with images written to a separate folder, then scans that folder, groups file names case‑insensitively, and reports any naming collisions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare temporary image file to embed in the workbook
                string tempImagePath = Path.Combine(Path.GetTempPath(), "sample_image.png");
                CreateSamplePng(tempImagePath);

                // Create a new workbook and add the image to the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Pictures.Add(0, 0, tempImagePath); // add image at cell A1

                // Configure HTML save options: export images as separate files (not Base64)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = false // ensure images are saved as files
                };

                // Define output HTML file path
                string outputHtml = Path.Combine(Directory.GetCurrentDirectory(), "output.html");

                // Save the workbook as HTML
                workbook.Save(outputHtml, htmlOptions);
                Console.WriteLine($"Workbook saved to HTML: {outputHtml}");

                // Determine the folder where Aspose.Cells stores the exported images
                // By default it creates a folder named "<htmlFileName>_files"
                string htmlDirectory = Path.GetDirectoryName(outputHtml);
                if (string.IsNullOrEmpty(htmlDirectory))
                {
                    Console.WriteLine("Unable to determine HTML directory.");
                    return;
                }

                string imageFolder = Path.Combine(
                    htmlDirectory,
                    Path.GetFileNameWithoutExtension(outputHtml) + "_files");

                if (Directory.Exists(imageFolder))
                {
                    // Get all image files generated for the HTML
                    string[] imageFiles = Directory.GetFiles(imageFolder);
                    Console.WriteLine($"Number of image files generated: {imageFiles.Length}");

                    // Verify that there are no duplicate file names (case‑insensitive)
                    var duplicateGroups = imageFiles
                        .Select(Path.GetFileName)
                        .GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
                        .Where(g => g.Count() > 1)
                        .Select(g => new { FileName = g.Key, Count = g.Count() })
                        .ToList();

                    if (duplicateGroups.Count == 0)
                    {
                        Console.WriteLine("No duplicate image files were created.");
                    }
                    else
                    {
                        Console.WriteLine("Duplicate image files detected:");
                        foreach (var dup in duplicateGroups)
                        {
                            Console.WriteLine($"  {dup.FileName} – occurrences: {dup.Count}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Image folder not found; no images were exported.");
                }

                // Clean up temporary image file
                if (File.Exists(tempImagePath))
                {
                    File.Delete(tempImagePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to create a simple 1x1 PNG image for testing
        private static void CreateSamplePng(string filePath)
        {
            // Minimal 1x1 red PNG (transparent background not required)
            byte[] pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
                0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
                0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
                0xAE,0x42,0x60,0x82
            };
            File.WriteAllBytes(filePath, pngBytes);
        }
    }
}
