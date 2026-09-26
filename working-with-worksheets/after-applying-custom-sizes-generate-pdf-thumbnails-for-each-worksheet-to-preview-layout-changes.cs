// Title: Apply custom column widths and row heights to every worksheet and generate PNG thumbnails using Aspose.Cells for .NET
// AI Prompts: Develop a C# console program that loops through all worksheets, changes column A to a width of 20 characters and row 1 to a height of 30 points, saves the modified workbook, and creates a PNG thumbnail for each sheet with Aspose.Cells SheetRender. | Implement a .NET solution that uses ImageOrPrintOptions with OnePagePerSheet to render each worksheet as a separate PNG preview after applying layout adjustments with Aspose.Cells.
// Common Searches: asp.net generate a preview image for every Excel worksheet after adjusting column sizes | c# Aspose.Cells produce PNG files for each sheet following row height changes | how to use SheetRender to output separate PNGs per worksheet in .NET | save workbook with custom column widths then create visual previews of all sheets
// Tags: column width adjustment Aspose.Cells C# | row height modification Aspose.Cells | SheetRender PNG export per worksheet | ImageOrPrintOptions OnePagePerSheet setting | worksheet thumbnail generation Aspose.Cells | apply layout changes workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing Excel file, iterates through each worksheet to set column A width to 20 characters and row 1 height to 30 points, saves the updated workbook, and then uses SheetRender with OnePagePerSheet to create a PNG thumbnail for every sheet in a designated folder.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string thumbnailFolder = "Thumbnails";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply custom column widths and row heights to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set column A (index 0) width to 20 characters
                sheet.Cells.SetColumnWidth(0, 20);
                // Set row 1 (index 0) height to 30 points
                sheet.Cells.SetRowHeight(0, 30);
            }

            // Ensure the thumbnail folder exists
            Directory.CreateDirectory(thumbnailFolder);

            // Configure image rendering options (PNG is default, one page per sheet)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Generate a PNG thumbnail for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    string thumbPath = Path.Combine(thumbnailFolder, $"{sheet.Name}_thumb.png");
                    SheetRender renderer = new SheetRender(sheet, imgOptions);
                    // Render the first (and only) page directly to a file
                    renderer.ToImage(0, thumbPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create thumbnail for sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Save the workbook with the applied size changes
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }

            Console.WriteLine("Processing completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
