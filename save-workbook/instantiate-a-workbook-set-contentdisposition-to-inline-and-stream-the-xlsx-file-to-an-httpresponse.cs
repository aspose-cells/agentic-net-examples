using System;
using System.IO;
using System.Net.Mime;                     // For ContentDisposition
using Aspose.Cells;                        // Core workbook classes

public static class WorkbookExportHelper
{
    /// <summary>
    /// Creates a simple workbook and streams the XLSX file to the supplied output stream.
    /// The caller can write the stream to an HTTP response, file, etc.
    /// </summary>
    /// <param name="outputStream">The stream to which the workbook will be saved.</param>
    public static void ExportWorkbook(Stream outputStream)
    {
        if (outputStream == null)
            throw new ArgumentNullException(nameof(outputStream));

        try
        {
            // 1. Instantiate a new workbook (default format is XLSX)
            var workbook = new Workbook();

            // 2. Add sample data (optional, just to demonstrate content)
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // 3. Prepare the ContentDisposition – set Inline = true so the browser opens the file
            var disposition = new ContentDisposition
            {
                Inline = true,               // Open in browser rather than force download
                FileName = "SampleWorkbook.xlsx"
            };

            // 4. The caller can use disposition.ToString() to set HTTP headers if needed.

            // 5. Save workbook directly to the provided stream in XLSX format
            workbook.Save(outputStream, SaveFormat.Xlsx);
        }
        catch (Exception)
        {
            // Rethrow after optional logging
            throw;
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a file stream to write the workbook to disk (adjust path as needed)
            string outputPath = "SampleWorkbook.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                WorkbookExportHelper.ExportWorkbook(fileStream);
            }

            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}