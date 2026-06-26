using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsDemo
{
    public class WorkbookExporter
    {
        // Exports a newly created workbook as an XLSX file to the provided stream.
        public void ExportWorkbook(Stream outputStream, string fileName = "Report.xlsx")
        {
            try
            {
                if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));

                // Create a new workbook (default format is XLSX)
                var workbook = new Workbook();

                // Add sample data to the first worksheet
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

                // Configure save options for the XLSX format
                var saveOptions = new OoxmlSaveOptions();

                // Save the workbook to the provided stream
                workbook.Save(outputStream, saveOptions);
            }
            catch (Exception ex)
            {
                // Wrap and rethrow to let the caller handle the failure
                throw new InvalidOperationException("Failed to export workbook.", ex);
            }
        }

        // Loads a workbook from a file if it exists; otherwise throws a descriptive exception.
        public Workbook LoadWorkbook(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    throw new ArgumentException("File path must be provided.", nameof(filePath));

                if (!File.Exists(filePath))
                    throw new FileNotFoundException($"The file '{filePath}' was not found.", filePath);

                return new Workbook(filePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to load workbook from '{filePath}'.", ex);
            }
        }
    }

    public static class Program
    {
        // Entry point required for compilation.
        public static void Main(string[] args)
        {
            try
            {
                var exporter = new WorkbookExporter();

                // Define output file path
                string outputPath = "Report.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Export workbook to file
                using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    exporter.ExportWorkbook(fileStream, Path.GetFileName(outputPath));
                }

                Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}