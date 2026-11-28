// Title: How to avoid creating duplicate image files when exporting an Excel workbook to HTML with ExportImagesAsBase64 = false using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export workbook to HTML without duplicate image files | C# save Excel as HTML with ExportImagesAsBase64 false and prevent image duplication | verify single image file generated for multiple picture inserts using Aspose.Cells | prevent duplicate PNG files when converting Excel to HTML in .NET | custom attached files folder handling after Aspose.Cells HTML export
// Tags: Aspose.Cells HTML export duplicate image prevention | ExportImagesAsBase64 false image handling | reuse same picture multiple cells Aspose.Cells | custom attached files directory Aspose.Cells HTML | detect duplicate exported images C#

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example creates a workbook, inserts the same PNG picture into three cells, and saves the workbook as HTML with ExportImagesAsBase64 set to false so images are written as separate files. It then moves the generated image files to a custom attached‑files folder and scans that folder to count image files and detect any filename duplicates, confirming that no duplicate image files are produced.
class Program
{
    static void Main()
    {
        // Prepare temporary working directories
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeCellsHtmlExport_" + Guid.NewGuid().ToString("N"));
        string attachedFilesDir = Path.Combine(workDir, "AttachedFiles");
        Directory.CreateDirectory(attachedFilesDir);

        // Paths for the generated files
        string htmlOutputPath = Path.Combine(workDir, "Workbook.html");
        string tempImagePath = Path.Combine(workDir, "tempImage.png");

        try
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and insert the same image multiple times
            // ------------------------------------------------------------
            // Create a simple 1x1 PNG image (red pixel) from a base64 string
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(tempImagePath, pngBytes);

            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Insert the same picture at three different locations using the file path
            ws.Pictures.Add(0, 0, tempImagePath);   // A1
            ws.Pictures.Add(5, 5, tempImagePath);   // F6
            ws.Pictures.Add(10, 2, tempImagePath);  // C11

            // ------------------------------------------------------------
            // 2. Configure HtmlSaveOptions with ExportImagesAsBase64 = false
            // ------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = false,               // images will be saved as separate files
                ExportActiveWorksheetOnly = true            // only the active sheet is exported
            };

            // ------------------------------------------------------------
            // 3. Save the workbook as HTML
            // ------------------------------------------------------------
            wb.Save(htmlOutputPath, htmlOptions);
            Console.WriteLine($"HTML saved to: {htmlOutputPath}");

            // Move exported image files to the desired attached files directory
            string defaultAttachedDir = Path.ChangeExtension(htmlOutputPath, null) + "_files";
            if (Directory.Exists(defaultAttachedDir))
            {
                foreach (string file in Directory.GetFiles(defaultAttachedDir))
                {
                    string destPath = Path.Combine(attachedFilesDir, Path.GetFileName(file));
                    File.Copy(file, destPath, true);
                }
                // Optionally delete the default folder
                Directory.Delete(defaultAttachedDir, true);
            }

            Console.WriteLine($"Attached files directory: {attachedFilesDir}");

            // ------------------------------------------------------------
            // 4. Verify that duplicate image files are NOT created
            // ------------------------------------------------------------
            if (!Directory.Exists(attachedFilesDir))
            {
                Console.WriteLine("Attached files directory does not exist – no images were exported.");
                return;
            }

            // Get all files that are images (common extensions)
            string[] imageExtensions = new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg" };
            var imageFiles = Directory.GetFiles(attachedFilesDir)
                                      .Where(f => imageExtensions.Contains(Path.GetExtension(f).ToLower()))
                                      .ToList();

            Console.WriteLine($"Number of image files exported: {imageFiles.Count}");

            // Group by file name (case‑insensitive) to detect duplicates
            var duplicateGroups = imageFiles
                .GroupBy(f => Path.GetFileName(f), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .ToList();

            if (duplicateGroups.Any())
            {
                Console.WriteLine("Duplicate image files were found:");
                foreach (var grp in duplicateGroups)
                {
                    Console.WriteLine($"  {grp.Key} – occurrences: {grp.Count()}");
                }
            }
            else
            {
                Console.WriteLine("Success: No duplicate image files were created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during processing:");
            Console.WriteLine(ex.Message);
        }
        finally
        {
            // Clean up temporary image file
            if (File.Exists(tempImagePath))
            {
                try { File.Delete(tempImagePath); } catch { }
            }

            // Optional: clean up temporary working directory (comment out if you want to inspect the output)
            // Directory.Delete(workDir, true);
        }
    }
}
