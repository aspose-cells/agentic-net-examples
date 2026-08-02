// Title: Aspose.Cells for .NET – Implement IFilePathProvider to Create Unique HTML Files for Worksheets with Duplicate or Sanitized Names
// Description: Shows how to build a custom IFilePathProvider that sanitizes worksheet titles with CellsHelper.CreateSafeSheetName, tracks case‑insensitive occurrences, and returns distinct file names (adding a numeric suffix for duplicates) when saving a workbook to HTML via HtmlSaveOptions.
// Keywords: Aspose.Cells | IFilePathProvider | unique HTML file name | duplicate worksheet names | CreateSafeSheetName | HtmlSaveOptions | export workbook to HTML | C# example | case‑insensitive sheet name handling | multiple HTML files per workbook
// Common Searches: Aspose.Cells custom IFilePathProvider example | generate unique HTML files for duplicate sheet names | how to sanitize worksheet names in Aspose.Cells | export each worksheet to separate HTML file .NET | unique file naming for Aspose.Cells HTML export
// Developer Intent: Create a custom IFilePathProvider that returns a unique file path for every worksheet, handling duplicate or sanitized titles during HTML export.
// Use Cases: Export a workbook with sheets named "Report*" and "Report?" to separate HTML files without overwriting. | Save workbooks that contain case‑insensitive duplicate sheet titles to distinct HTML files automatically. | Integrate UniqueFilePathProvider into HtmlSaveOptions to produce one HTML file per worksheet with incremental suffixes.
// AI Prompts: Write a C# class implementing IFilePathProvider that adds an incremental suffix to duplicate sanitized worksheet names for HTML export with Aspose.Cells. | Provide a step‑by‑step tutorial on using UniqueFilePathProvider with HtmlSaveOptions to generate separate HTML files for each worksheet. | Explain the role of CellsHelper.CreateSafeSheetName and how to combine it with a dictionary to ensure unique file names during workbook export.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom file path provider that ensures each worksheet gets a unique file name
    // even when worksheet titles are duplicated (case‑insensitive or after sanitizing).
    // Shows how to build a custom IFilePathProvider that sanitizes worksheet titles with CellsHelper.CreateSafeSheetName, tracks case‑insensitive occurrences, and returns distinct file names (adding a numeric suffix for duplicates) when saving a workbook to HTML via HtmlSaveOptions.
    internal class UniqueFilePathProvider : IFilePathProvider
    {
        // Tracks how many times a sanitized sheet name has been encountered.
        private readonly Dictionary<string, int> _nameCounters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public string GetFullName(string sheetName)
        {
            // Convert the original sheet name to a safe file‑system name.
            // This removes invalid characters and truncates to Excel's 31‑character limit.
            string safeName = CellsHelper.CreateSafeSheetName(sheetName);

            // Determine the occurrence count for this safe name.
            if (_nameCounters.TryGetValue(safeName, out int count))
            {
                // Increment the counter for subsequent duplicates.
                count++;
                _nameCounters[safeName] = count;
            }
            else
            {
                // First occurrence of this name.
                count = 0;
                _nameCounters[safeName] = count;
            }

            // Build a unique file name. The first occurrence gets the base name,
            // subsequent duplicates receive a numeric suffix.
            string fileName = count == 0 ? $"{safeName}.html" : $"{safeName}_{count}.html";

            // Return the full path (here we assume the current directory; callers can prepend a folder if needed).
            return fileName;
        }
    }

    public class IFilePathProviderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and populate it with data.
                Workbook workbook = new Workbook();

                // First worksheet – default name "Sheet1".
                Worksheet ws1 = workbook.Worksheets[0];
                // Use a name containing illegal characters; it will be sanitized before assignment.
                ws1.Name = CellsHelper.CreateSafeSheetName("Report*"); // Becomes "Report"
                ws1.Cells["A1"].PutValue("First sheet");

                // Add a second worksheet with a name that sanitizes to the same safe name as the first.
                int idx2 = workbook.Worksheets.Add();
                Worksheet ws2 = workbook.Worksheets[idx2];
                ws2.Name = CellsHelper.CreateSafeSheetName("Report?"); // Also becomes "Report"
                ws2.Cells["A1"].PutValue("Second sheet");

                // Add a third worksheet with a distinct safe name.
                int idx3 = workbook.Worksheets.Add();
                Worksheet ws3 = workbook.Worksheets[idx3];
                ws3.Name = "Summary";
                ws3.Cells["A1"].PutValue("Third sheet");

                // Configure HTML save options to use the custom file path provider.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = false, // Export all worksheets.
                    FilePathProvider = new UniqueFilePathProvider()
                };

                // Save the workbook; each worksheet will be written to a separate HTML file
                // with unique names generated by UniqueFilePathProvider.
                workbook.Save("CombinedOutput.html", saveOptions);

                Console.WriteLine("Workbook saved with custom unique file names for each worksheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static void Main()
        {
            IFilePathProviderDemo.Run();
        }
    }
}
