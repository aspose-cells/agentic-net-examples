// Title: Save each worksheet of an Excel workbook as a separate PNG image using Aspose.Cells for .NET
// Common Searches: C# code to export every sheet in an .xlsx file to individual PNG files with Aspose.Cells | How to loop through worksheets and generate PNG images using Aspose.Cells .NET | Aspose.Cells render each worksheet to PNG one page per sheet example | Convert Excel worksheets to PNG images in C# with error handling | Save Excel sheet as PNG using SheetRender and ImageOrPrintOptions in Aspose.Cells
// Tags: Aspose.Cells worksheet to PNG conversion | C# SheetRender export PNG per sheet | ImageOrPrintOptions OnePagePerSheet setting | Iterate workbook worksheets Aspose.Cells | Save Excel sheet as image .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook with Aspose.Cells, iterates through all worksheets, configures ImageOrPrintOptions for PNG with OnePagePerSheet enabled, renders each sheet via SheetRender, and saves the output as "{SheetName}.png" while handling missing files and rendering exceptions.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Verify that the file exists before attempting to load
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"File not found: {sourcePath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // LoadOptions can be used to specify a password if needed
            LoadOptions loadOptions = new LoadOptions();
            // Uncomment and set the password if the workbook is protected
            // loadOptions.Password = "yourPassword";

            workbook = new Workbook(sourcePath, loadOptions);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
            return;
        }

        // Iterate through each worksheet and export to PNG
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // Configure image options for PNG format
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Png,
                OnePagePerSheet = true // render the whole sheet on one page
            };

            try
            {
                // Render the worksheet to an image
                SheetRender renderer = new SheetRender(sheet, imgOptions);

                // Define the output PNG file name (using the sheet name)
                string pngFile = $"{sheet.Name}.png";

                // Save the first (and only) page as PNG
                renderer.ToImage(0, pngFile);
                Console.WriteLine($"Saved sheet \"{sheet.Name}\" as {pngFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to render sheet \"{sheet.Name}\": {ex.Message}");
            }
        }
    }
}
