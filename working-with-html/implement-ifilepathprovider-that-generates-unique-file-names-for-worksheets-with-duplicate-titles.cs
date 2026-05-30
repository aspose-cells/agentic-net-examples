using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom implementation of IFilePathProvider that ensures unique file names
    // for worksheets with duplicate (case‑insensitive) titles.
    internal class UniqueFilePathProvider : IFilePathProvider
    {
        // Tracks how many times a safe sheet name has been used.
        private readonly Dictionary<string, int> _nameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public string GetFullName(string sheetName)
        {
            // Convert the original sheet name to a safe Excel sheet name.
            string safeName = CellsHelper.CreateSafeSheetName(sheetName);

            // Determine the current count for this safe name.
            if (_nameCounts.TryGetValue(safeName, out int count))
            {
                // Increment count and generate a new unique file name.
                count++;
                _nameCounts[safeName] = count;
                return $"{safeName}_{count}.html";
            }
            else
            {
                // First occurrence – store count = 0 and return the base name.
                _nameCounts[safeName] = 0;
                return $"{safeName}.html";
            }
        }
    }

    public class IFilePathProviderDuplicateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a workbook with duplicate sheet titles.
                Workbook workbook = new Workbook();

                // First sheet (default name is "Sheet1").
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Report";

                // Add a second sheet and intentionally give it the same title.
                Worksheet sheet2 = workbook.Worksheets.Add("Report");
                sheet2.Cells["A1"].PutValue("Second sheet with duplicate title");

                // Add a third sheet with a title that differs only by case.
                Worksheet sheet3 = workbook.Worksheets.Add("report");
                sheet3.Cells["A1"].PutValue("Third sheet with case‑insensitive duplicate title");

                // Configure HTML save options to use the custom file path provider.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export each worksheet to a separate HTML file.
                    ExportActiveWorksheetOnly = false,
                    // Assign the custom provider.
                    FilePathProvider = new UniqueFilePathProvider()
                };

                // Define output file name.
                string outputFile = "WorkbookWithUniqueHtmlFiles.html";

                // Save the workbook; each worksheet will be written to a uniquely named HTML file.
                workbook.Save(outputFile, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during demo execution: {ex.Message}");
            }
        }
    }

    // Entry point for the console application.
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                IFilePathProviderDuplicateDemo.Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}