// Title: Refresh Linked Pictures After Updating Source Images with Aspose.Cells for .NET (C#)
// Description: A C# example that overwrites an external PNG, loads the Excel workbook, iterates through Worksheet.Pictures, reassigns SourceFullName for each linked picture to force a refresh, and saves the file so the latest image appears in the workbook.
// Keywords: Aspose.Cells linked picture refresh | C# update external image Excel | force reload linked picture Aspose | SourceFullName refresh .NET | Excel workbook external image update | Aspose.Cells picture IsLink | GitHub Aspose.Cells example
// Common Searches: how to refresh linked picture in Excel using Aspose.Cells | Aspose.Cells C# reload external image | force linked picture to update after source file change | Aspose.Cells refresh pictures programmatically | update external PNG in Excel workbook .NET
// Developer Intent: Programmatically reload linked pictures in an Excel file after the source image on disk has been modified.
// Use Cases: Regenerate reports that reference external graphics and need the newest version of each image. | Batch‑process multiple workbooks where linked images may be replaced on the file system. | Automate a workflow that swaps logo files and ensures all linked logos in existing spreadsheets reflect the change.
// AI Prompts: Write C# code with Aspose.Cells that refreshes all linked pictures after the source image files are edited. | Show how to detect linked pictures (IsLink) and force them to reload using SourceFullName. | Explain error handling for missing source image files when refreshing linked pictures in an Excel workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A C# example that overwrites an external PNG, loads the Excel workbook, iterates through Worksheet.Pictures, reassigns SourceFullName for each linked picture to force a refresh, and saves the file so the latest image appears in the workbook.
class RefreshLinkedPictures
{
    static void Main()
    {
        try
        {
            // Paths for the workbook and the linked image file
            string workbookPath = "LinkedPictureDemo.xlsx";
            string imagePath = "linked_image.png";
            string newImagePath = "new_image_content.png";

            // -------------------------------------------------
            // Step 0: Verify required files exist
            // -------------------------------------------------
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Ensure the linked image file exists (create an empty placeholder if missing)
            if (!File.Exists(imagePath))
            {
                File.WriteAllBytes(imagePath, new byte[0]);
            }

            // -------------------------------------------------
            // Step 1: Modify the source image on the file system
            // -------------------------------------------------
            // Overwrite the linked image with new content if the new image file is present.
            if (File.Exists(newImagePath))
            {
                byte[] newImageData = File.ReadAllBytes(newImagePath);
                File.WriteAllBytes(imagePath, newImageData);
                Console.WriteLine("Linked image updated with new content.");
            }
            else
            {
                Console.WriteLine($"New image file not found: {newImagePath}. Skipping image update.");
            }

            // -------------------------------------------------
            // Step 2: Load the workbook that contains the linked picture
            // -------------------------------------------------
            Workbook workbook = new Workbook(workbookPath);
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Step 3: Locate linked pictures and refresh them
            // -------------------------------------------------
            foreach (Picture pic in sheet.Pictures)
            {
                if (pic.IsLink)
                {
                    // Reassign the same source path to force a refresh of the linked image
                    string currentSource = pic.SourceFullName;
                    pic.SourceFullName = currentSource;
                }
            }

            // -------------------------------------------------
            // Step 4: Save the workbook (the linked picture will now reflect the updated image)
            // -------------------------------------------------
            string outputPath = "LinkedPictureDemo_Refreshed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
