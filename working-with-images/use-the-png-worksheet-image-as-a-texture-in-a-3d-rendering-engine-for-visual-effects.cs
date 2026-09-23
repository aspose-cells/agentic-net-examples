// Title: Export the first worksheet of an Excel workbook to a PNG file with Aspose.Cells for use as a 3D texture in C#
// AI Prompts: Generate C# code that loads a workbook, renders the first worksheet to a PNG image using Aspose.Cells, and returns the image path. | Show how to load the exported PNG as a texture in a Unity or OpenGL C# project after creating it with Aspose.Cells. | Create a robust C# snippet that verifies the Excel file exists, exports the worksheet to PNG with Aspose.Cells, handles exceptions, and prepares the image for a graphics engine.
// Common Searches: how to save an Excel worksheet as a PNG image using Aspose.Cells in .NET | C# convert first sheet of Excel to PNG for Unity texture | Aspose.Cells SheetRender export worksheet to PNG for OpenGL texture mapping | sample code to render Excel sheet to PNG and load it as a texture in a C# 3D engine
// Tags: Aspose.Cells PNG export of worksheet | SheetRender PNG generation in C# | Unity texture from Excel sheet PNG | OpenGL PNG texture loading C# | exception handling for Excel export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads Sample.xlsx, extracts the first worksheet, and uses Aspose.Cells' SheetRender to export it as Worksheet.png. It includes error handling for missing files and placeholders for integrating the PNG as a texture in OpenGL or Unity 3D rendering pipelines.
class WorksheetTextureDemo
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "Sample.xlsx";

        // Verify that the Excel file exists to avoid FileNotFoundException
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(excelPath);

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image export options – PNG format will be inferred from the output file extension
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Optional: set resolution, background, etc.
                // HorizontalResolution = 300,
                // VerticalResolution = 300
            };

            // Render the worksheet to an image (page index 0)
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);

            // Export the worksheet as PNG directly to a file
            string outputPath = "Worksheet.png";
            sheetRender.ToImage(0, outputPath);
            Console.WriteLine($"Worksheet exported to \"{outputPath}\".");

            // If OpenGL texture creation is required, it can be added here
            // using an appropriate OpenGL wrapper library.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        // Keep console open
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
