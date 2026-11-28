// Title: Export Excel worksheets to separate HTML files while preserving inter‑sheet links using Aspose.Cells IFilePathProvider in C#
// Common Searches: c# Aspose.Cells export each worksheet to its own html file preserving sheet links | how to implement IFilePathProvider for custom worksheet html filenames in Aspose.Cells | save workbook as multiple html pages with clickable links between sheets using Aspose.Cells | Aspose.Cells HtmlSaveOptions keep full‑path hyperlinks when exporting to html | generate safe file names for worksheet html export with Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions export multiple worksheets | IFilePathProvider custom file naming Aspose.Cells | preserve inter‑worksheet hyperlinks HTML export | generate worksheet‑specific html files C# | safe worksheet name to file name conversion

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorksheetsToHtml
{
    // Implements IFilePathProvider to control the file name of each worksheet HTML file.
    // The GetFilePath method receives the worksheet name and its zero‑based index.
    // It must return a unique file path (relative or absolute) for the HTML file.
    // The example shows how to implement a custom IFilePathProvider that creates safe, sheet‑name‑based HTML file names, configure HtmlSaveOptions to export all worksheets, enable full‑path links so hyperlinks between sheets remain functional, and save the workbook as a main index HTML file plus one HTML file per worksheet using Aspose.Cells for .NET.
    public class CustomFilePathProvider : IFilePathProvider
    {
        // Return a simple file name based on the worksheet name.
        // You can customize the naming scheme as needed.
        public string GetFilePath(string sheetName, int sheetIndex)
        {
            // Ensure the sheet name is a valid file name.
            string safeName = MakeFileSystemSafe(sheetName);
            return $"{safeName}.html";
        }

        // Required by IFilePathProvider – returns the full path for a given file name.
        public string GetFullName(string fileName)
        {
            // Resolve to an absolute path; this helps Aspose.Cells locate the file.
            return Path.GetFullPath(fileName);
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
            // Path to the source workbook. Adjust as necessary.
            const string inputFile = "input.xlsx";

            // Verify that the input file exists before proceeding.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputFile);

                // Configure HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Export all worksheets (default is false, but set explicitly for clarity).
                htmlOptions.ExportActiveWorksheetOnly = false;

                // Preserve hyperlinks that point to other worksheets.
                // When true, the generated HTML will contain links like "Sheet2.html#A1".
                htmlOptions.IsFullPathLink = true;

                // Assign our custom file‑path provider so each worksheet gets its own HTML file.
                htmlOptions.FilePathProvider = new CustomFilePathProvider();

                // The main HTML file (index) name. It will contain links to each worksheet file.
                const string mainHtmlFile = "WorkbookIndex.html";

                // Save the workbook as HTML. This will create the main file plus one HTML file per worksheet.
                workbook.Save(mainHtmlFile, htmlOptions);

                Console.WriteLine("Export completed successfully.");
                Console.WriteLine($"Main index file: {Path.GetFullPath(mainHtmlFile)}");
                Console.WriteLine("Generated worksheet files:");

                // List the generated files using the same naming logic as the provider.
                var provider = new CustomFilePathProvider();
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    string sheetFile = provider.GetFilePath(sheet.Name, sheet.Index);
                    Console.WriteLine($" - {Path.GetFullPath(sheetFile)}");
                }
            }
            catch (Exception ex)
            {
                // Any exception (e.g., loading, saving) is reported to the console.
                Console.WriteLine("An error occurred during export:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
