// Title: Export a Workbook to TIFF on a UNC Network Share with Aspose.Cells (C#)
// Description: Demonstrates how to create a Workbook, set ImageOrPrintOptions, render a worksheet with SheetRender, and save the output as a multi‑page TIFF directly to a UNC network share using the ToTiff method, including basic error handling.
// Keywords: Aspose.Cells | C# | .NET | TIFF export | UNC path | network share | SheetRender | ToTiff | ImageOrPrintOptions | save TIFF to shared folder | export workbook to TIFF
// Common Searches: Aspose.Cells export TIFF to UNC path | C# save TIFF on network share | SheetRender ToTiff example | How to write TIFF file to shared folder using Aspose.Cells | Export Excel as TIFF to remote folder .NET
// Developer Intent: Save the generated TIFF file directly to a network share path.
// Use Cases: Automatically archive daily reports as TIFF images on a file server for compliance. | Provide a centralized print queue by storing multi‑page TIFFs in a shared folder accessible to multiple workstations. | Create snapshot images of spreadsheets for offline review and store them on a network drive.
// AI Prompts: Generate C# code that uses Aspose.Cells to render a workbook to a multi‑page TIFF and saves it to a UNC network share with exception handling. | Explain how ImageOrPrintOptions.OnePagePerSheet affects TIFF output when exporting to a remote folder. | Show how to modify the sample to loop through all worksheets and save each as an individual TIFF file in a shared directory.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsNetworkTiffDemo
{
    // Demonstrates how to create a Workbook, set ImageOrPrintOptions, render a worksheet with SheetRender, and save the output as a multi‑page TIFF directly to a UNC network share using the ToTiff method, including basic error handling.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample content
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Network Share TIFF Demo");

            // Configure rendering options (optional, but demonstrates usage)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true // render each sheet as a separate page
            };

            // Initialize the sheet renderer (lifecycle rule: create)
            SheetRender sheetRenderer = new SheetRender(worksheet, renderOptions);

            // Define the UNC network share path where the TIFF will be saved
            string networkSharePath = @"\\MyServer\SharedFolder\DemoOutput.tiff";

            try
            {
                // Render the worksheet to a TIFF file directly on the network share
                // (lifecycle rule: save using ToTiff(string))
                sheetRenderer.ToTiff(networkSharePath);
                Console.WriteLine($"TIFF successfully saved to network share: {networkSharePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving TIFF to network share: {ex.Message}");
            }
        }
    }
}
