// Title: Using IFilePathProvider to define custom HTML file names for each worksheet when exporting a workbook with Aspose.Cells in C#
// Common Searches: c# aspocells set custom html file name for each worksheet using IFilePathProvider | how to export Excel workbook to separate html files per sheet with Aspose.Cells | Aspose.Cells HtmlSaveOptions FilePathProvider example for safe file names | custom output directory for Aspose.Cells HTML export in .NET | prevent invalid characters in generated html file names Aspose.Cells
// Tags: IFilePathProvider implementation Aspose.Cells | HTML export per worksheet Aspose.Cells | custom file path provider C# | safe file name generation .NET | HtmlSaveOptions FilePathProvider usage | export workbook to multiple html files

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Implements IFilePathProvider to control the full file path of each worksheet HTML file.
    // Demonstrates how to implement IFilePathProvider to generate safe, directory‑specific HTML file names for each worksheet during Aspose.Cells HTML export, including workbook loading, HtmlSaveOptions configuration, custom provider assignment, saving, and listing of generated files.
    public class CustomFilePathProvider : IFilePathProvider
    {
        private readonly string _outputDirectory;

        public CustomFilePathProvider(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        // Returns the full path (including folder) for a generated HTML file.
        public string GetFullName(string fileName)
        {
            // Ensure the file name is safe for the file system.
            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            string safeName = MakeFileSystemSafe(nameWithoutExt);
            string extension = Path.GetExtension(fileName);

            // Combine with the output directory.
            return Path.Combine(_outputDirectory, $"{safeName}{extension}");
        }

        // Helper to replace characters that are illegal in file names.
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Prepare input workbook.
            // -----------------------------------------------------------------
            string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
            string inputPath = Path.Combine(dataDir, "SampleWorkbook.xlsx");

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to load workbook:");
                Console.WriteLine(ex.Message);
                return;
            }

            // -----------------------------------------------------------------
            // 2. Configure HTML export options.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false,   // Export all worksheets.
                ExportImagesAsBase64 = true          // Embed images as Base64.
            };

            // -----------------------------------------------------------------
            // 3. Define output folder and base file name.
            // -----------------------------------------------------------------
            string outputDir = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
            Directory.CreateDirectory(outputDir);

            // Provide the custom file path provider so each worksheet gets a predictable name.
            htmlOptions.FilePathProvider = new CustomFilePathProvider(outputDir);

            // The base file name is used for the first worksheet; others are handled by the provider.
            string baseHtmlFile = Path.Combine(outputDir, "Workbook.html");

            // -----------------------------------------------------------------
            // 4. Save the workbook as HTML.
            // -----------------------------------------------------------------
            try
            {
                // Save using the overload that accepts HtmlSaveOptions.
                workbook.Save(baseHtmlFile, htmlOptions);
                Console.WriteLine("Workbook successfully exported to HTML.");
                Console.WriteLine($"Base file: {baseHtmlFile}");
                Console.WriteLine($"All worksheet files are located in: {outputDir}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during HTML export:");
                Console.WriteLine(ex.Message);
            }

            // -----------------------------------------------------------------
            // 5. List generated HTML files (simple integrity check).
            // -----------------------------------------------------------------
            Console.WriteLine("\nGenerated HTML files:");
            foreach (string file in Directory.GetFiles(outputDir, "*.html"))
            {
                Console.WriteLine($" - {Path.GetFileName(file)}");
            }
        }
    }
}
