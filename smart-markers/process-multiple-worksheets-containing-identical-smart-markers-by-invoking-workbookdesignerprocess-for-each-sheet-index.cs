// Title: Use Aspose.Cells WorkbookDesigner in C# to process identical smart markers on every worksheet of an Excel workbook
// AI Prompts: Write a C# console program that loads an Excel file, creates a WorkbookDesigner, calls Process to fill identical smart markers on all sheets, and saves the workbook. | Show how to validate input and output paths, ensure the output directory exists, and handle exceptions when processing smart markers with Aspose.Cells. | Demonstrate using Aspose.Cells WorkbookDesigner to apply the same smart‑marker data source across multiple worksheets without manually iterating each sheet.
// Common Searches: Aspose.Cells C# process smart markers on all worksheets in one call | how to apply the same smart marker data to every sheet using WorkbookDesigner | C# example for bulk smart marker processing with Aspose.Cells workbookdesigner | save processed Excel after smart marker replacement in C# Aspose.Cells | error handling for WorkbookDesigner.Process when input file missing
// Tags: Aspose.Cells WorkbookDesigner process all worksheets | C# bulk smart marker replacement Excel | smart markers identical across multiple sheets | validate input file path Aspose.Cells | create output directory before saving workbook C# | exception handling WorkbookDesigner.Process

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a WorkbookDesigner, processes identical smart markers on every worksheet with a single Process call, ensures the output folder exists, saves the updated file, and includes robust path validation and error handling.
class Program
{
    static void Main(string[] args)
    {
        // Input and output file paths can be passed via command‑line arguments or hard‑coded.
        string inputPath = args.Length > 0 ? args[0] : "{InputFile}";
        string outputPath = args.Length > 1 ? args[1] : "{OutputFile}";

        try
        {
            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // Create a designer to process smart markers.
            var designer = new WorkbookDesigner(workbook);

            // Process smart markers for all worksheets.
            designer.Process();

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the processed workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed successfully and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any errors.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
