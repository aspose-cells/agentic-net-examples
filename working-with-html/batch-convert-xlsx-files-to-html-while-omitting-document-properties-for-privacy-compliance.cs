// Title: Batch convert XLSX to HTML without metadata using Aspose.Cells for .NET
// Description: An example C# console app that scans an input folder, loads each .xlsx workbook with Aspose.Cells, and saves it as HTML in an output folder. HtmlSaveOptions are set to ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties = false, ensuring the generated HTML contains no Excel metadata, meeting privacy‑compliance requirements.
// Keywords: Aspose.Cells | C# batch XLSX to HTML | convert Excel to HTML .NET | exclude document properties | HtmlSaveOptions | privacy compliant Excel conversion | bulk Excel HTML export | ExportDocumentProperties false | ExportWorkbookProperties false | ExportWorksheetProperties false
// Common Searches: C# batch convert XLSX files to HTML Aspose.Cells | How to hide Excel metadata when saving as HTML | Aspose.Cells HtmlSaveOptions omit document properties | Convert multiple Excel workbooks to HTML without properties | Privacy safe Excel to HTML conversion .NET
// Developer Intent: Convert multiple Excel workbooks to HTML while stripping all document, workbook, and worksheet properties.
// Use Cases: Publishing confidential Excel reports on a website without exposing metadata | Automating nightly HTML export of financial spreadsheets for intranet dashboards while ensuring data privacy | Building a file‑sanitization service that removes Excel properties before sharing | Generating static HTML documentation from Excel templates in a CI pipeline
// AI Prompts: Write C# code using Aspose.Cells to batch convert all .xlsx files in a folder to .html with ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties disabled. | Describe how each HtmlSaveOptions flag influences the resulting HTML and how to verify that no metadata is present. | Suggest ways to add progress logging, error handling, and CSV reporting to the batch conversion script. | Explain performance considerations for large workbooks during bulk HTML export with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// An example C# console app that scans an input folder, loads each .xlsx workbook with Aspose.Cells, and saves it as HTML in an output folder. HtmlSaveOptions are set to ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties = false, ensuring the generated HTML contains no Excel metadata, meeting privacy‑compliance requirements.
class BatchXlsxToHtml
{
    static void Main()
    {
        // Folder containing source XLSX files
        string sourceFolder = "InputXlsx";

        // Folder where HTML files will be written
        string outputFolder = "OutputHtml";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the source folder
        foreach (string xlsxPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            // Safety check – the file should exist
            if (!File.Exists(xlsxPath))
            {
                Console.WriteLine($"File not found (skipped): {xlsxPath}");
                continue;
            }

            try
            {
                // Load the workbook from the XLSX file
                Workbook workbook = new Workbook(xlsxPath);

                // Configure HTML save options to exclude all document‑related properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportDocumentProperties = false,   // omit built‑in document properties
                    ExportWorkbookProperties = false,   // omit workbook‑level properties
                    ExportWorksheetProperties = false   // omit worksheet‑level properties
                };

                // Determine the output HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Converted: {xlsxPath} -> {htmlPath}");
            }
            catch (Exception ex)
            {
                // Log any errors but continue processing other files
                Console.WriteLine($"Error processing '{xlsxPath}': {ex.Message}");
            }
        }
    }
}
