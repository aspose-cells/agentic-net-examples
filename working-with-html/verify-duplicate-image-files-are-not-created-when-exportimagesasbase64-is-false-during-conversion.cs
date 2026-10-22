// Title: Check for duplicate image files after converting an Excel workbook to HTML with ExportImagesAsBase64 disabled using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook, save it as HTML with HtmlSaveOptions.ExportImagesAsBase64 = false, then enumerate the generated image folder and compute MD5 hashes to flag any duplicate images. | Write a C# routine that scans the <workbook>_files directory produced by Aspose.Cells HTML export, compares each file's hash, and logs the names of duplicate image files. | Adjust HtmlSaveOptions to ensure images are saved as separate files, run the conversion, and verify that no identical image files exist in the output folder.
// Common Searches: how to detect duplicate images after exporting Excel to HTML with Aspose.Cells .NET | Aspose.Cells HTML export ExportImagesAsBase64 false duplicate image files check | C# verify that Aspose.Cells does not create duplicate image files during HTML conversion | compare MD5 hashes of images extracted by Aspose.Cells HTML save options
// Tags: htmlsaveoptions exportimagesasbase64 false | aspocells duplicate image detection html export | excel to html image extraction .net | md5 hash comparison for image files c# | validate extracted image folder aspocells

using System;
using System.IO;
using System.Collections.Generic;
using System.Security.Cryptography;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel workbook, saves it as HTML with ExportImagesAsBase64 disabled, then scans the resulting <html>_files directory, computes MD5 hashes for each extracted image, and reports whether any duplicate image files were created.
class Program
{
    static void Main()
    {
        try
        {
            // Verify input workbook exists
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Configure HTML save options (export images as separate files)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = false
            };

            // Prepare output paths
            string htmlOutputPath = Path.Combine("output", "output.html");
            string htmlDir = Path.GetDirectoryName(htmlOutputPath);
            if (!string.IsNullOrEmpty(htmlDir))
                Directory.CreateDirectory(htmlDir);

            // Save the workbook as HTML
            try
            {
                workbook.Save(htmlOutputPath, saveOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save HTML: {ex.Message}");
                return;
            }

            // Determine the folder where Aspose.Cells extracts images (default: <htmlFileName>_files)
            string imageFolder = Path.Combine(
                htmlDir ?? string.Empty,
                Path.GetFileNameWithoutExtension(htmlOutputPath) + "_files");

            // Verify that duplicate image files were not created
            if (Directory.Exists(imageFolder))
            {
                string[] imageFiles = Directory.GetFiles(imageFolder);
                var hashSet = new HashSet<string>();
                bool duplicateFound = false;

                foreach (string filePath in imageFiles)
                {
                    try
                    {
                        using (FileStream stream = File.OpenRead(filePath))
                        using (MD5 md5 = MD5.Create())
                        {
                            string hash = BitConverter.ToString(md5.ComputeHash(stream));
                            if (!hashSet.Add(hash))
                            {
                                duplicateFound = true;
                                Console.WriteLine($"Duplicate image detected: {Path.GetFileName(filePath)}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing image '{filePath}': {ex.Message}");
                    }
                }

                if (!duplicateFound)
                    Console.WriteLine("No duplicate image files were created.");
            }
            else
            {
                Console.WriteLine("Image folder not found; no images to verify.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
