using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchXlsxToHtml
{
    class Program
    {
        static void Main()
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\InputXlsx";
            // Folder where the generated HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Prepare load options for XLSX files
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Iterate over each .xlsx file in the source folder
            foreach (string xlsxPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Configure HTML save options to omit all document‑related properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
                htmlOptions.ExportDocumentProperties = false;   // Do not export built‑in document properties
                htmlOptions.ExportWorkbookProperties = false;   // Do not export workbook properties
                htmlOptions.ExportWorksheetProperties = false; // Do not export worksheet properties

                // Determine the output HTML file path
                string htmlPath = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(xlsxPath) + ".html");

                // Convert the XLSX file to HTML using the utility method with options
                ConversionUtility.Convert(xlsxPath, loadOptions, htmlPath, htmlOptions);

                Console.WriteLine($"Converted '{xlsxPath}' to '{htmlPath}' (properties omitted).");
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}