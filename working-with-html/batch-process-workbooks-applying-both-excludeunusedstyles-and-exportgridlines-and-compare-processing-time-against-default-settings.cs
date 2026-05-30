using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // For HtmlSaveOptions

class BatchProcessWorkbooks
{
    static void Main()
    {
        // Folder containing source Excel files
        string sourceFolder = @"C:\InputWorkbooks";
        // Folder for output HTML files with custom options
        string customOutputFolder = @"C:\OutputWorkbooks\Custom";
        // Folder for output HTML files with default options
        string defaultOutputFolder = @"C:\OutputWorkbooks\Default";

        Directory.CreateDirectory(customOutputFolder);
        Directory.CreateDirectory(defaultOutputFolder);

        // Process each .xlsx file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            // Load workbook (lifecycle: create/load)
            Workbook wb = new Workbook(filePath);

            // -------------------- Custom options (ExcludeUnusedStyles + ExportGridLines) --------------------
            var customOptions = new HtmlSaveOptions();
            customOptions.ExcludeUnusedStyles = true;   // Exclude unused styles (default is true)
            customOptions.ExportGridLines = true;      // Export grid lines (default is false)

            // Optional: remove unused styles from the workbook before saving
            wb.RemoveUnusedStyles();

            string customOutputPath = Path.Combine(customOutputFolder,
                Path.GetFileNameWithoutExtension(filePath) + "_custom.html");

            var swCustom = Stopwatch.StartNew();
            // Save with custom HTML options (lifecycle: save)
            wb.Save(customOutputPath, customOptions);
            swCustom.Stop();

            // -------------------- Default options --------------------
            // Reload workbook to avoid any side‑effects from previous save
            wb = new Workbook(filePath);

            string defaultOutputPath = Path.Combine(defaultOutputFolder,
                Path.GetFileNameWithoutExtension(filePath) + "_default.html");

            var swDefault = Stopwatch.StartNew();
            // Save with default settings (no HtmlSaveOptions supplied)
            wb.Save(defaultOutputPath);
            swDefault.Stop();

            // -------------------- Report timing --------------------
            Console.WriteLine($"File: {Path.GetFileName(filePath)}");
            Console.WriteLine($"  Custom (ExcludeUnusedStyles + ExportGridLines): {swCustom.ElapsedMilliseconds} ms");
            Console.WriteLine($"  Default: {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine();
        }
    }
}