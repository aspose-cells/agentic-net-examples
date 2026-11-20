// Title: C# batch conversion of XLSX files to PDF with page margins and orientation preserved using Aspose.Cells
// Common Searches: C# Aspose.Cells convert multiple Excel files in a folder to PDF preserving page setup | How to keep worksheet margins and orientation when exporting XLSX to PDF with Aspose.Cells | Batch process all .xlsx files to PDF in .NET while retaining page layout | Command line utility for bulk Excel to PDF conversion using Aspose.Cells SaveFormat.Pdf | Convert a directory of Excel workbooks to PDF without losing page orientation in C#
// Tags: batch xlsx to pdf conversion Aspose.Cells | preserve worksheet page setup during PDF export | Aspose.Cells SaveFormat.Pdf with existing margins | C# bulk Excel to PDF conversion folder | retain page orientation in Excel PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// The C# console app validates input and output folders, enumerates all .xlsx files in the specified directory, loads each workbook with Aspose.Cells (which automatically retains page margins and orientation), and saves each workbook as a PDF using SaveFormat.Pdf. It logs each conversion and handles errors, enabling efficient bulk Excel‑to‑PDF conversion while preserving original page layout settings.
class BatchXlsxToPdfConverter
{
    static void Main(string[] args)
    {
        // Validate arguments: args[0] = input folder, args[1] = output folder
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: BatchXlsxToPdfConverter <inputFolder> <outputFolder>");
            return;
        }

        string inputFolder = args[0];
        string outputFolder = args[1];

        // Ensure input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Create output folder if it does not exist
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Get all .xlsx files in the input folder (non-recursive)
        string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string xlsxPath in xlsxFiles)
        {
            try
            {
                // Load the workbook (preserves all worksheet settings, including page margins and orientation)
                Workbook workbook = new Workbook(xlsxPath);

                // Optional: verify that page setup settings are retained (no changes needed)
                // foreach (Worksheet sheet in workbook.Worksheets)
                // {
                //     PageSetup setup = sheet.PageSetup;
                //     // setup.Orientation, setup.LeftMargin, etc., are already set in the source file
                // }

                // Build the output PDF file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                // Save as PDF; Aspose.Cells respects the existing PageSetup (margins, orientation)
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(pdfPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file '{xlsxPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
