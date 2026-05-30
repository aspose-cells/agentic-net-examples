using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchXlsxToHtml
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\InputXlsx";

            // Folder where the generated HTML files will be saved
            string destinationFolder = @"C:\OutputHtml";

            // Ensure the destination folder exists
            Directory.CreateDirectory(destinationFolder);

            // Get all XLSX files in the source folder
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

            foreach (string xlsxPath in xlsxFiles)
            {
                // Build the output HTML file path (same name, .html extension)
                string htmlFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".html";
                string htmlPath = Path.Combine(destinationFolder, htmlFileName);

                // Configure HTML save options: disable CSS generation (use only inline styles)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    DisableCss = true,               // Inline styles only, no external CSS
                    ExcludeUnusedStyles = true       // Exclude unused styles to reduce size (default)
                };

                // Convert the XLSX file to HTML using the ConversionUtility rule
                // LoadOptions is set to null to use default loading behavior
                ConversionUtility.Convert(xlsxPath, null, htmlPath, htmlOptions);

                Console.WriteLine($"Converted '{xlsxPath}' to '{htmlPath}' with CSS disabled.");
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}