// Title: Verify that HTML export with ExportImagesAsBase64=false does not generate duplicate image files using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, saves it to HTML with ExportImagesAsBase64 set to false, and then scans the exported image folder to detect duplicate images by comparing SHA‑256 hashes. | Create a method that computes a hash for each image file produced by Aspose.Cells during HTML conversion and throws an exception when two files share the same hash value. | Show how to programmatically locate the default image subfolder created by Aspose.Cells when saving to HTML and verify that each exported image is unique.
// Common Searches: Aspose.Cells export to HTML without base64 images duplicate file detection | C# verify unique image files after saving workbook as HTML using Aspose.Cells | how to prevent duplicate image files when ExportImagesAsBase64 is false in Aspose.Cells | check for identical exported images in Aspose.Cells HTML conversion .NET
// Tags: HTML export separate image files Aspose.Cells | ExportImagesAsBase64 false handling | duplicate image detection C# | SHA256 hash comparison for images .NET | Aspose.Cells generated image folder pattern

using System;
using System.IO;
using System.Collections.Generic;
using System.Security.Cryptography;
using Aspose.Cells;

// The sample loads an Excel workbook, saves it as HTML with ExportImagesAsBase64 disabled so images are written to a separate folder, then iterates over the folder, computes SHA‑256 hashes for each image, and raises an error if any two images share the same hash, confirming that no duplicate image files were created.
class VerifyNoDuplicateImages
{
    static void Main()
    {
        try
        {
            // Verify that the input workbook exists
            string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input.xlsx");
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Input workbook not found.", inputPath);

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare HTML save options with ExportImagesAsBase64 set to false
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.ExportImagesAsBase64 = false; // Images will be saved as separate files

            // Determine the folder where Aspose.Cells will export images (default naming convention)
            string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.html");
            string imageFolder = Path.Combine(
                Path.GetDirectoryName(htmlPath) ?? string.Empty,
                Path.GetFileNameWithoutExtension(htmlPath) + "_files");

            // Ensure the image folder exists (Aspose.Cells will create it if needed)
            Directory.CreateDirectory(imageFolder);

            // Save the workbook to HTML (images will be exported to the folder)
            workbook.Save(htmlPath, saveOptions);

            // Verify that no duplicate image files were created
            // Duplicate is defined as files having identical content (hash)
            var hashSet = new HashSet<string>();
            foreach (string filePath in Directory.GetFiles(imageFolder))
            {
                // Compute SHA256 hash of the file content
                using (FileStream stream = File.OpenRead(filePath))
                using (SHA256 sha = SHA256.Create())
                {
                    byte[] hashBytes = sha.ComputeHash(stream);
                    string hashString = BitConverter.ToString(hashBytes).Replace("-", "");

                    // If the hash already exists, a duplicate image file was created
                    if (!hashSet.Add(hashString))
                    {
                        throw new InvalidOperationException(
                            $"Duplicate image detected: {Path.GetFileName(filePath)} has the same content as another exported image.");
                    }
                }
            }

            Console.WriteLine("Verification completed: No duplicate image files were created.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
