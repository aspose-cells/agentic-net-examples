// Title: Refresh Linked Pictures in Excel with Aspose.Cells for .NET after Source Image Changes
// Description: Shows how to force external linked images to reload in an Excel workbook using Aspose.Cells for .NET. The sample walks through each worksheet, detects linked Picture objects, reassigns their SourceFullName to trigger a refresh, and saves the workbook so updated graphics are displayed.
// Keywords: Aspose.Cells refresh linked pictures | C# reload external images Excel | picture.IsLink Aspose.Cells | SourceFullName refresh image | update linked picture programmatically | Excel workbook external image update | force picture reload .NET | linked picture refresh code
// Common Searches: how to refresh linked images in Excel using Aspose.Cells | force linked picture to reload after source file change C# | update external picture references in a workbook programmatically | Aspose.Cells refresh linked picture after editing PNG | reload Excel linked images with .NET code
// Developer Intent: Reload linked pictures so that modifications to source image files are reflected in the saved Excel workbook.
// Use Cases: After editing logo files on disk, run the routine to ensure the workbook shows the newest version before distribution. | Batch‑process multiple reports to refresh all linked graphics (charts, logos, watermarks) in one step. | Integrate the refresh logic into an automated document generation pipeline where source assets may change between runs.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets and refreshes every linked picture after the source image files have been modified. | Explain why reassigning Picture.SourceFullName forces a linked picture to reload in Aspose.Cells and suggest alternative methods, if any.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to force external linked images to reload in an Excel workbook using Aspose.Cells for .NET. The sample walks through each worksheet, detects linked Picture objects, reassigns their SourceFullName to trigger a refresh, and saves the workbook so updated graphics are displayed.
class RefreshLinkedPictures
{
    static void Main()
    {
        // Load the workbook that contains linked pictures
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through each picture in the worksheet
            foreach (Picture picture in sheet.Pictures)
            {
                // Process only linked pictures
                if (picture.IsLink)
                {
                    // Reassign the source path to force the picture to reload the image
                    string sourcePath = picture.SourceFullName;
                    picture.SourceFullName = sourcePath;
                }
            }
        }

        // Save the workbook with refreshed linked images
        workbook.Save("output.xlsx");
    }
}
