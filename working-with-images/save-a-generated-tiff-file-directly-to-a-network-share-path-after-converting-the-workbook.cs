using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SaveTiffToNetworkShare
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Aspose.Cells TIFF Export Demo");

            // Configure rendering options (optional settings)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true // Render the whole sheet on a single page
            };

            // Initialize the sheet renderer with the worksheet and options
            SheetRender sheetRenderer = new SheetRender(worksheet, renderOptions);

            // Define the network share path where the TIFF will be saved
            string networkSharePath = @"\\Server\Share\Folder\output.tiff";

            // Ensure the target directory exists; create it if necessary
            string targetDir = Path.GetDirectoryName(networkSharePath);
            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            // Render the worksheet directly to a TIFF file on the network share
            sheetRenderer.ToTiff(networkSharePath);

            Console.WriteLine($"TIFF file successfully saved to: {networkSharePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}