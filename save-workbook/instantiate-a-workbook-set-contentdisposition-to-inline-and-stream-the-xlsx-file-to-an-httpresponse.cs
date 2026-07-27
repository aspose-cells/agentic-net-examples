using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class WorkbookExporter
    {
        // Exports a workbook to the provided stream (e.g., HttpResponse.Body) with inline disposition.
        public void Export(Stream outputStream, string fileName = "sample.xlsx")
        {
            if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));

            try
            {
                // 1. Create a new workbook and add sample data.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // 2. Save the workbook directly to the stream in Xlsx format.
                workbook.Save(outputStream, SaveFormat.Xlsx);
                // Note: Setting the Content‑Disposition header (inline; filename=…) is the caller's responsibility.
            }
            catch (Exception ex)
            {
                // Wrap and rethrow for caller handling.
                throw new InvalidOperationException("Failed to export workbook.", ex);
            }
        }

        // Optional helper to export directly to a file, ensuring the path exists.
        public void ExportToFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            try
            {
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // 1. Create workbook with sample data.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // 2. Save to file in Xlsx format.
                workbook.Save(filePath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to export workbook to '{filePath}'.", ex);
            }
        }
    }

    // Simple console entry point for demonstration/testing.
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var exporter = new WorkbookExporter();

                // Example: export to a file.
                string outputPath = Path.Combine(Environment.CurrentDirectory, "output", "sample.xlsx");
                exporter.ExportToFile(outputPath);
                Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}