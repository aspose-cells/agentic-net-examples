// Title: Refresh Linked Pictures in an Excel Workbook with Aspose.Cells for .NET
// Description: Loads a workbook, scans each worksheet for linked Picture objects, forces each image to reload by reassigning its SourceFullName, and saves the updated file. Ideal for synchronizing images after external changes on web servers.
// Keywords: Aspose.Cells linked picture refresh | C# reload Excel images | update picture source Aspose.Cells | force picture reload .NET | Excel workbook linked images update | global image synchronization Excel
// Common Searches: how to refresh linked pictures in Excel using Aspose.Cells | programmatically reload external images in a workbook .NET | reset picture source path Aspose.Cells C# | update Excel linked images after web server change
// Developer Intent: Programmatically refresh every linked image in a workbook after its source file has changed.
// Use Cases: Ensure the latest corporate logo appears in generated reports by refreshing linked images before export. | Synchronize chart snapshots stored as external PNG files on a web server with a shared Excel template. | Automatically update product photos in a sales catalog workbook when the source images are edited.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets and resets Picture.SourceFullName to refresh linked images. | Explain error handling strategies when refreshing linked pictures in an Excel file using Aspose.Cells for .NET. | Show how to batch‑process multiple workbooks to refresh external image links with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, scans each worksheet for linked Picture objects, forces each image to reload by reassigning its SourceFullName, and saves the updated file. Ideal for synchronizing images after external changes on web servers.
class RefreshLinkedPictures
{
    static void Main()
    {
        // Load the workbook that contains linked pictures
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through all pictures in the current worksheet
            foreach (Picture picture in sheet.Pictures)
            {
                // Process only linked pictures
                if (picture.IsLink)
                {
                    // Reassign the same source path to force the picture to reload
                    string sourcePath = picture.SourceFullName;
                    picture.SourceFullName = sourcePath;
                }
            }
        }

        // Save the workbook after refreshing the linked pictures
        workbook.Save("output.xlsx");
    }
}
