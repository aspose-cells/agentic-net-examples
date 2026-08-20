// Title: C# – Batch Convert XLSX to HTML with Aspose.Cells, Excluding All Metadata
// Description: Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, and saves it as .html using HtmlSaveOptions that disable ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties, ensuring the generated HTML contains no spreadsheet metadata.
// Keywords: Aspose.Cells | C# batch convert XLSX to HTML | exclude document properties | privacy compliant Excel to HTML | HtmlSaveOptions ExportDocumentProperties false | remove workbook metadata | ExportWorkbookProperties false | ExportWorksheetProperties false | GDPR Excel conversion | metadata stripping HTML export
// Common Searches: batch convert xlsx to html aspose.cells c# | asp.net hide excel metadata when exporting to html | htmlsaveoptions exportdocumentproperties false example | convert folder of excel files to html without properties | privacy compliant excel to html conversion c#
// Developer Intent: Convert multiple Excel workbooks to HTML while stripping all built‑in document, workbook, and worksheet properties.
// Use Cases: Create web‑ready reports from a bulk of Excel files without exposing any spreadsheet metadata. | Automate GDPR‑compliant conversion of user‑uploaded spreadsheets to HTML in a SaaS application. | Archive financial Excel data as static HTML for audit trails while ensuring no hidden properties remain.
// AI Prompts: Generate C# code that batch converts every .xlsx in a directory to .html with Aspose.Cells, setting ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties to false. | Explain how to also remove custom document properties when exporting workbooks to HTML using Aspose.Cells. | Provide a snippet to log each conversion (source, destination, status) to a CSV file while keeping all metadata exclusion settings active.

using System;
using System.IO;
using Aspose.Cells;

// Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, and saves it as .html using HtmlSaveOptions that disable ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties, ensuring the generated HTML contains no spreadsheet metadata.
class BatchConvertXlsxToHtml
{
    static void Main()
    {
        // Folder containing source XLSX files
        string sourceFolder = "InputXlsx";
        // Folder where HTML files will be saved
        string outputFolder = "OutputHtml";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder \"{sourceFolder}\" not found. Please ensure the folder exists and contains XLSX files.");
            return;
        }

        // Create output directory if it does not exist
        Directory.CreateDirectory(outputFolder);

        // Retrieve all XLSX files in the source folder
        string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string xlsxPath in xlsxFiles)
        {
            // Ensure the file still exists before processing
            if (!File.Exists(xlsxPath))
            {
                Console.WriteLine($"File not found: {xlsxPath}");
                continue;
            }

            try
            {
                // Load the workbook from the XLSX file
                Workbook workbook = new Workbook(xlsxPath);

                // Set HTML save options to exclude all document-related properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportDocumentProperties = false,   // Omit built‑in document properties
                    ExportWorkbookProperties = false,   // Omit workbook properties
                    ExportWorksheetProperties = false   // Omit worksheet properties
                };

                // Build the output HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Converted: {xlsxPath} -> {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file \"{xlsxPath}\": {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion of XLSX to HTML completed.");
    }
}
