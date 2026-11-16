// Title: Convert multiple XLSX workbooks to HTML with Aspose.Cells for .NET while suppressing document properties
// Common Searches: C# Aspose.Cells batch convert all .xlsx files in a folder to HTML without metadata | How to export Excel to HTML using Aspose.Cells and hide document properties | Aspose.Cells HtmlSaveOptions ExportDocumentProperties false example in .NET | Programmatically convert a directory of XLSX files to HTML in C# | Privacy‑compliant Excel to HTML conversion using Aspose.Cells
// Tags: Aspose.Cells batch XLSX to HTML conversion | HtmlSaveOptions ExportDocumentProperties false | C# convert Excel workbooks to HTML | privacy compliant Excel HTML export | process multiple Excel files with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace BatchConversion
{
    // The sample scans a given input folder for .xlsx files, loads each workbook with Aspose.Cells, sets HtmlSaveOptions.ExportDocumentProperties to false to omit metadata, and saves the result as .html files in a specified output directory, creating the directory if needed.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Provide the folder containing XLSX files and the folder where HTML files will be saved.
            string inputFolder = @"C:\InputXlsx";
            string outputFolder = @"C:\OutputHtml";

            BatchConvertXlsxToHtml(inputFolder, outputFolder);
        }

        /// <param name="inputFolder">Folder containing source XLSX files.</param>
        /// <param name="outputFolder">Folder where HTML files will be written.</param>
        static void BatchConvertXlsxToHtml(string inputFolder, string outputFolder)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Get all .xlsx files in the input directory (non-recursive).
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Load the workbook from the XLSX file.
                Workbook workbook = new Workbook(xlsxPath);

                // Configure HTML save options to exclude document properties.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportDocumentProperties = false   // Omit document properties for privacy.
                };

                // Determine the output HTML file path.
                string htmlFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML using the configured options.
                workbook.Save(htmlPath, htmlOptions);
            }
        }
    }
}
