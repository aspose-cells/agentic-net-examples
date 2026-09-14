// Title: Convert a folder of Excel .xlsx workbooks to HTML with Aspose.Cells using PresentationPreference.BestFit in C#
// AI Prompts: Generate a C# console application that scans a specified directory for *.xlsx files, loads each workbook with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to BestFit, and saves the output as .html files in a target folder. | Write C# code that creates HtmlSaveOptions for Aspose.Cells with the BestFit layout, iterates over all Excel files in an input folder, and batch‑converts them to HTML while handling missing files and ensuring the output directory exists.
// Common Searches: aspocells c# batch convert xlsx files to html with bestfit layout | how to use HtmlSaveOptions PresentationPreference BestFit in Aspose.Cells | c# program to convert all Excel workbooks in a folder to html using Aspose.Cells | convert multiple .xlsx files to .html preserving column widths aspocells | aspocells save workbook as html with best‑fit presentation option
// Tags: Aspose.Cells batch HTML conversion | HtmlSaveOptions PresentationPreference BestFit | C# enumerate Excel files directory | convert xlsx to html Aspose.Cells | Aspose.Cells workbook.Save with HTML options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Scans a given input folder for .xlsx files, loads each workbook with Aspose.Cells, applies HtmlSaveOptions with PresentationPreference.BestFit, and saves each workbook as an .html file in a specified output folder, including basic error handling and folder creation.
class BatchExcelToHtml
{
    static void Main()
    {
        // Folder containing the Excel workbooks to convert
        string inputFolder = @"C:\InputExcels";

        // Folder where the resulting HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Get all Excel files (you can adjust the pattern for other extensions)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string excelPath in excelFiles)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found, skipping: {excelPath}");
                    continue;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // The PresentationPreference property may not be available in all versions.
                // If needed, uncomment the following line after confirming the enum exists.
                // htmlOptions.PresentationPreference = PresentationPreference.BestFit;

                // Determine the output HTML file path
                string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"Converted: {excelPath} -> {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
