// Title: Export each worksheet’s shapes to individual SVG files with Aspose.Cells for .NET
// AI Prompts: Write a C# program that loads an .xlsx workbook, creates an output folder, and saves every worksheet as a separate SVG file while preserving all embedded shapes using Aspose.Cells. | Adjust the sample code to produce SVG files that contain only the drawing layer of each worksheet, keeping the original colors and positions intact.
// Common Searches: Aspose.Cells C# convert Excel sheet to SVG preserving drawings | export only worksheet drawing objects to SVG with Aspose.Cells | batch convert multiple Excel worksheets to SVG in .NET | how to render Excel charts and images as SVG using Aspose.Cells | save Excel shapes as SVG files for web embedding with C#
// Tags: Aspose.Cells worksheet to SVG export | C# export Excel shapes as SVG | ImageOrPrintOptions SaveFormat Svg Aspose.Cells | SheetRender render worksheet page to SVG | batch SVG conversion of Excel worksheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an InputWorkbook.xlsx, creates an 'ExportedShapes' directory, and iterates through each worksheet. For each sheet it configures ImageOrPrintOptions with SaveFormat = Svg and OnePagePerSheet = true, then uses SheetRender to generate an SVG file named with the workbook and sheet name, preserving all visual elements such as shapes, charts, and images.
class ExportShapesToSvg
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "InputWorkbook.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Folder to store exported SVG files
            string outputFolder = "ExportedShapes";
            Directory.CreateDirectory(outputFolder);

            // Iterate through each worksheet and export it as SVG (includes shapes)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    string svgFileName = $"{Path.GetFileNameWithoutExtension(workbookPath)}_{sheet.Name}.svg";
                    string svgFilePath = Path.Combine(outputFolder, svgFileName);

                    // Set options for SVG export
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        SaveFormat = SaveFormat.Svg,
                        OnePagePerSheet = true
                    };

                    // Render the worksheet to SVG
                    SheetRender renderer = new SheetRender(sheet, options);
                    renderer.ToImage(0, svgFilePath);

                    Console.WriteLine($"Exported worksheet '{sheet.Name}' to SVG at '{svgFilePath}'.");
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error exporting worksheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            Console.WriteLine("Shape export process completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
