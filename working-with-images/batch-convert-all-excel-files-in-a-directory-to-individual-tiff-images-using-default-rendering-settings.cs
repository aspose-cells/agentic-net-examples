// Title: Batch convert Excel files to individual TIFF images per worksheet with Aspose.Cells (C#)
// Description: Scans a folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and renders every worksheet to a separate TIFF file using default ImageOrPrintOptions. Files are saved with a sanitized "Workbook_Sheet.tiff" name in a target directory.
// Keywords: Aspose.Cells batch TIFF conversion | C# export Excel worksheets as images | convert multiple Excel files to TIFF | render Excel sheet to TIFF default settings | automated Excel to image conversion
// Common Searches: C# batch convert Excel worksheets to TIFF | Aspose.Cells export each sheet as TIFF file | how to render all Excel files in a folder to TIFF images | save Excel worksheets as separate TIFF pictures | default image options Aspose.Cells TIFF output
// Developer Intent: Automatically generate a TIFF image for every worksheet in every Excel workbook located in a specified directory, using Aspose.Cells with default rendering options.
// Use Cases: Archive financial workbooks as high‑resolution TIFFs for regulatory compliance. | Create printable snapshots of dashboards from dozens of Excel reports. | Produce image assets for web galleries or documentation without manual export.
// AI Prompts: Generate C# code that converts all Excel files in a directory to PNG images per worksheet using Aspose.Cells. | Show how to add custom DPI and page margins to the TIFF conversion sample. | Provide error‑handling patterns for missing input folders, corrupted workbooks, and unsupported file types.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Scans a folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and renders every worksheet to a separate TIFF file using default ImageOrPrintOptions. Files are saved with a sanitized "Workbook_Sheet.tiff" name in a target directory.
class BatchExcelToTiff
{
    static void Main()
    {
        // Directory containing source Excel files
        string inputDir = "InputExcels";

        // Directory where TIFF images will be saved
        string outputDir = "OutputTiffs";
        Directory.CreateDirectory(outputDir);

        // Get all Excel files (xls, xlsx, xlsm) in the input directory
        string[] excelFiles = Directory.GetFiles(inputDir, "*.*", SearchOption.TopDirectoryOnly)
            .Where(f => f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        foreach (string filePath in excelFiles)
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Process each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Use default image rendering options
                ImageOrPrintOptions options = new ImageOrPrintOptions();

                // Create a SheetRender for the current worksheet
                SheetRender renderer = new SheetRender(sheet, options);

                // Build a safe output file name: <WorkbookName>_<SheetName>.tiff
                string safeSheetName = GetSafeFileName(sheet.Name);
                string tiffFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_{safeSheetName}.tiff";
                string tiffPath = Path.Combine(outputDir, tiffFileName);

                // Render the whole worksheet to a TIFF file
                renderer.ToTiff(tiffPath);

                // Release resources used by the renderer
                renderer.Dispose();
            }
        }
    }

    // Replaces characters that are invalid in file names with an underscore
    static string GetSafeFileName(string name)
    {
        char[] invalidChars = Path.GetInvalidFileNameChars();
        foreach (char c in invalidChars)
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
