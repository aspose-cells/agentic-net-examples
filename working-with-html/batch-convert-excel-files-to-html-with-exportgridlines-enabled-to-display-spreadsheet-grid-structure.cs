// Title: Batch Convert Excel Files to HTML with Gridlines Using Aspose.Cells (.NET)
// Description: Scans a folder, loads each .xls, .xlsx, .xlsb, .xlsm, or .csv workbook with the appropriate LoadFormat, and uses Aspose.Cells ConversionUtility together with HtmlSaveOptions (ExportGridLines = true) to generate an HTML file for every workbook in a target directory, while handling missing files and conversion errors.
// Keywords: Aspose.Cells | C# | .NET | batch Excel to HTML conversion | ExportGridLines | HtmlSaveOptions | ConversionUtility | load format | XLSX to HTML | CSV to HTML | automated spreadsheet export | web preview of Excel
// Common Searches: Aspose.Cells batch convert Excel to HTML with gridlines | C# convert folder of .xlsx files to HTML preserving cell borders | How to export Excel files as HTML using ExportGridLines | Convert multiple CSV and XLSX files to HTML with Aspose.Cells | Sample code for Aspose.Cells HTMLSaveOptions ExportGridLines
// Developer Intent: Automatically transform every supported Excel workbook in a directory into an HTML page that shows the original gridlines.
// Use Cases: Create web‑ready HTML reports from a collection of financial spreadsheets while keeping cell borders visible. | Provide instant HTML previews of uploaded Excel or CSV files in a web portal without requiring Office installations. | Schedule nightly jobs that archive a folder of workbooks as static HTML pages for documentation or compliance purposes.
// AI Prompts: Show how to add a custom CSS file to HtmlSaveOptions while keeping ExportGridLines enabled. | Modify the batch conversion to skip hidden worksheets during HTML export. | Replace console logging with CSV logging of conversion results using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Scans a folder, loads each .xls, .xlsx, .xlsb, .xlsm, or .csv workbook with the appropriate LoadFormat, and uses Aspose.Cells ConversionUtility together with HtmlSaveOptions (ExportGridLines = true) to generate an HTML file for every workbook in a target directory, while handling missing files and conversion errors.
class BatchExcelToHtml
{
    static void Main()
    {
        // Folder containing source Excel files
        string inputFolder = "InputExcels";

        // Folder where HTML files will be saved
        string outputFolder = "OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify the input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder '{inputFolder}' does not exist.");
            return;
        }

        // Get all files in the input folder
        string[] files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string sourcePath in files)
        {
            // Skip if the file does not exist (safety check)
            if (!File.Exists(sourcePath))
                continue;

            string ext = Path.GetExtension(sourcePath).ToLowerInvariant();

            // Process only supported Excel formats
            if (ext == ".xls" || ext == ".xlsx" || ext == ".xlsb" || ext == ".xlsm" || ext == ".csv")
            {
                // Determine the appropriate load format
                LoadFormat loadFormat = GetLoadFormat(ext);
                LoadOptions loadOptions = new LoadOptions(loadFormat);

                // Configure HTML save options with gridlines enabled
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true
                };

                // Destination HTML file path
                string destPath = Path.Combine(outputFolder,
                    Path.GetFileNameWithoutExtension(sourcePath) + ".html");

                try
                {
                    // Perform the conversion using Aspose.Cells ConversionUtility
                    ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);
                    Console.WriteLine($"Converted '{sourcePath}' to '{destPath}' with gridlines.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to convert '{sourcePath}': {ex.Message}");
                }
            }
        }
    }

    // Maps file extensions to Aspose.Cells LoadFormat values
    static LoadFormat GetLoadFormat(string extension)
    {
        switch (extension)
        {
            case ".xls":  return LoadFormat.Excel97To2003;
            case ".xlsx": // .xlsm files are also loaded as Xlsx format
            case ".xlsm": return LoadFormat.Xlsx;
            case ".xlsb": return LoadFormat.Xlsb;
            case ".csv":  return LoadFormat.Csv;
            default:      return LoadFormat.Auto;
        }
    }
}
