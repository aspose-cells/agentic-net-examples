// Title: Export an existing workbook to XLSX and deliver it as an attachment in an ASP.NET HTTP response using Aspose.Cells (C#)
// Common Searches: how to return an Aspose.Cells workbook as a downloadable XLSX file in ASP.NET MVC | Aspose.Cells C# export workbook to MemoryStream and set Content-Disposition attachment | send generated Excel file as attachment using Aspose.Cells in Web API | save workbook to stream and trigger file download with Aspose.Cells in .NET Core | C# Aspose.Cells write XLSX to HttpResponse output stream
// Tags: export workbook to MemoryStream Aspose.Cells | XLSX attachment response Aspose.Cells | ASP.NET Core Excel file download Aspose.Cells | save workbook as XLSX C# Aspose.Cells | set Content-Disposition for Excel download | write Excel to HttpResponse stream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads a workbook with Aspose.Cells, validates inputs, saves it directly to a provided stream in XLSX format, resets the stream position if possible, and can be used to send the file as an attachment in an ASP.NET HTTP response.
    public class WorkbookExporter
    {
        // Exports an existing workbook as an XLSX to the provided output stream.
        public void ExportWorkbook(Stream outputStream, string sourceFilePath, string suggestedFileName = "YourWorkbook.xlsx")
        {
            if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));
            if (string.IsNullOrEmpty(sourceFilePath)) throw new ArgumentNullException(nameof(sourceFilePath));

            // Verify that the source workbook file exists to avoid FileNotFoundException.
            if (!File.Exists(sourceFilePath))
                throw new FileNotFoundException($"Workbook file not found: {sourceFilePath}");

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(sourceFilePath);

                // Write the workbook directly to the provided stream in XLSX format.
                workbook.Save(outputStream, SaveFormat.Xlsx);

                // Reset stream position for callers that may read from the beginning.
                if (outputStream.CanSeek)
                    outputStream.Position = 0;
            }
            catch (Exception)
            {
                // Rethrow to let caller handle.
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Define path to source workbook (adjust as needed).
            string sourcePath = @"C:\Path\To\YourWorkbook.xlsx";

            // If source file does not exist, create a simple workbook for demonstration.
            if (!File.Exists(sourcePath))
            {
                try
                {
                    Workbook wb = new Workbook();
                    wb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    // Ensure directory exists.
                    Directory.CreateDirectory(Path.GetDirectoryName(sourcePath));
                    wb.Save(sourcePath, SaveFormat.Xlsx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                    return;
                }
            }

            // Export the workbook to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                try
                {
                    WorkbookExporter exporter = new WorkbookExporter();
                    exporter.ExportWorkbook(ms, sourcePath);

                    // For demonstration, write the exported bytes to a file.
                    string outputPath = Path.Combine(Path.GetDirectoryName(sourcePath), "ExportedWorkbook.xlsx");
                    File.WriteAllBytes(outputPath, ms.ToArray());
                    Console.WriteLine($"Workbook exported successfully to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during export: {ex.Message}");
                }
            }
        }
    }
}
