// Title: Generate a multi‑page TIFF from only the 'Summary' and 'Data' worksheets using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells render only selected Excel sheets to a multi‑page TIFF | How to export specific worksheets (Summary, Data) as TIFF with Aspose.Cells | Create a TIFF file from particular sheets in a workbook using Aspose.Cells .NET | Aspose.Cells C# generate multi‑page TIFF with LZW compression for chosen worksheets | Copy selected worksheets to new workbook before TIFF conversion using Aspose.Cells
// Tags: WorkbookRender to TIFF Aspose.Cells | ImageOrPrintOptions LZW compression | one page per sheet rendering | select worksheets for image export | multi-page TIFF generation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The C# program loads an Excel workbook, copies only the "Summary" and "Data" worksheets into a new workbook, and uses Aspose.Cells' WorkbookRender with ImageOrPrintOptions (LZW compression, 300 dpi, OnePagePerSheet) to create a multi‑page TIFF file. It includes error handling for missing files or sheets and writes the TIFF to a specified output path.
class Program
{
    static void Main(string[] args)
    {
        // Input and output file paths.
        // If arguments are supplied, use them; otherwise fall back to default names.
        string sourcePath = args.Length > 0 ? args[0] : "source.xlsx";
        string outputPath = args.Length > 1 ? args[1] : "SelectedSheets.tiff";

        // Names of worksheets to include in the TIFF.
        string[] sheetsToExport = { "Summary", "Data" };

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook.
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook that will contain only the selected sheets.
            Workbook exportWorkbook = new Workbook();

            // Remove the default empty sheet that a new Workbook contains.
            exportWorkbook.Worksheets.Clear();

            // Copy each required sheet from the source workbook.
            foreach (string sheetName in sheetsToExport)
            {
                Worksheet sourceSheet = null;
                try
                {
                    // Attempt to retrieve the worksheet by name.
                    sourceSheet = sourceWorkbook.Worksheets[sheetName];
                }
                catch
                {
                    // Worksheet not found; keep sourceSheet null.
                }

                if (sourceSheet == null)
                {
                    Console.WriteLine($"Warning: Worksheet \"{sheetName}\" not found in source workbook.");
                    continue;
                }

                // Add a copy of the worksheet to the export workbook using the name overload.
                exportWorkbook.Worksheets.AddCopy(sourceSheet.Name);
                Console.WriteLine($"Added worksheet \"{sheetName}\" to export workbook.");
            }

            if (exportWorkbook.Worksheets.Count == 0)
            {
                Console.WriteLine("No matching worksheets were found. TIFF will not be created.");
                return;
            }

            // Configure image rendering options for TIFF output.
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Tiff,
                TiffCompression = TiffCompression.CompressionLZW,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                OnePagePerSheet = true
            };

            // Render the workbook (containing only the selected sheets) to a multi‑page TIFF file.
            WorkbookRender renderer = new WorkbookRender(exportWorkbook, renderOptions);
            renderer.ToImage(outputPath);

            Console.WriteLine($"TIFF file created successfully at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the TIFF:");
            Console.WriteLine(ex.Message);
        }
    }
}
