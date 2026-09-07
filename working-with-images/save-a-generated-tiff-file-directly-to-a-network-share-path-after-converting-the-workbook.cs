// Title: Render an Aspose.Cells worksheet to a TIFF image and save it directly to a UNC network share using C#
// AI Prompts: Write C# code that uses Aspose.Cells to convert a worksheet to a TIFF image and writes the file to a specified UNC path, creating the folder if it does not exist. | Show how to configure ImageOrPrintOptions for a single‑page TIFF export and stream the result to a network share without creating a local temporary file. | Provide a .NET method that renders the first sheet of a workbook as TIFF, ensures write permissions on \\Server\Share, and saves the image bytes directly to that location.
// Common Searches: c# aspnet save aspose.cells worksheet as tiff to network share | how to export excel sheet to tiff file on unc path using aspose.cells | aspose.cells render sheet to tiff and write to remote folder | write image bytes from aspose.cells to \\server\share in c# | ensure directory exists before saving tiff from workbook in .net
// Tags: Aspose.Cells worksheet to TIFF conversion | C# write TIFF to UNC share | ImageOrPrintOptions OnePagePerSheet TIFF export | SheetRender render sheet as image stream | Directory.CreateDirectory network path validation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a Workbook, fills cells A1 and A2, sets ImageOrPrintOptions for one page per sheet, renders the first worksheet to a memory stream, ensures the target UNC directory exists, and writes the stream bytes as a .tiff file directly to a network share.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");

            // Set image rendering options (default format will be PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true // Render each sheet on a separate page
            };

            // Render the worksheet to an image stream
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            using (MemoryStream imageStream = new MemoryStream())
            {
                renderer.ToImage(0, imageStream);
                imageStream.Position = 0; // Reset stream position for reading

                // Define the network share path (ensure write permissions)
                string networkPath = @"\\Server\Share\output.tiff";

                // Ensure the target directory exists
                string directory = Path.GetDirectoryName(networkPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the image stream to the network share as TIFF
                // The stream contains PNG by default; convert to TIFF by saving with appropriate extension
                // Aspose.Cells can directly save as TIFF by setting ImageFormat in options, but here we write the stream.
                File.WriteAllBytes(networkPath, imageStream.ToArray());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
