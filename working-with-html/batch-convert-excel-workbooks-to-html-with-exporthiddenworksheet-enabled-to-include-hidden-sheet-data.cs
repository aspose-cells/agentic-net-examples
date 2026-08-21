// Title: C# Batch Convert Excel Workbooks to HTML Including Hidden Sheets – Aspose.Cells Example
// Description: A C# .NET script that scans a folder, loads each .xls/.xlsx/.xlsm file with Aspose.Cells, sets HtmlSaveOptions.ExportHiddenWorksheet = true, and saves the workbooks as HTML files, preserving hidden worksheet data.
// Keywords: Aspose.Cells batch HTML conversion | ExportHiddenWorksheet C# | convert Excel to HTML .NET | hidden worksheets to HTML | C# folder Excel to HTML example | Aspose.Cells HtmlSaveOptions | batch Excel to HTML code | Aspose.Cells hidden sheet export
// Common Searches: how to batch convert Excel files to HTML with hidden sheets using Aspose.Cells | Aspose.Cells ExportHiddenWorksheet example C# | convert multiple xls xlsx files to HTML preserving hidden worksheets | C# code to export hidden worksheets to HTML with Aspose.Cells | Aspose.Cells HTML export hidden worksheets batch processing
// Developer Intent: Automatically convert every Excel workbook in a directory to HTML while including the content of hidden worksheets.
// Use Cases: Publish a collection of Excel reports to an intranet, ensuring hidden calculation sheets are visible in the HTML view. | Create archival HTML snapshots of spreadsheets for compliance, capturing all hidden data. | Generate web‑ready documentation from batch‑processed Excel files without manually opening each workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch convert all Excel files in a folder to HTML with ExportHiddenWorksheet enabled. | Show an Aspose.Cells .NET example that scans a directory, loads .xls/.xlsx/.xlsm files, and saves each as a single HTML file while preserving hidden worksheets. | Provide a step‑by‑step guide for configuring HtmlSaveOptions in Aspose.Cells to include hidden worksheets during bulk HTML export.

using System;
using System.IO;
using Aspose.Cells;

// A C# .NET script that scans a folder, loads each .xls/.xlsx/.xlsm file with Aspose.Cells, sets HtmlSaveOptions.ExportHiddenWorksheet = true, and saves the workbooks as HTML files, preserving hidden worksheet data.
class BatchExcelToHtml
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputExcel";

        // Folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (xls, xlsx, xlsm) from the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                continue; // Skip non‑Excel files

            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Set HTML save options to include hidden worksheets
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = true // Export content of hidden sheets
            };

            // Determine the output HTML file path (same name, .html extension)
            string outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(filePath) + ".html");

            // Save the workbook as HTML using the configured options
            workbook.Save(outputFile, htmlOptions);

            Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML.");
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
