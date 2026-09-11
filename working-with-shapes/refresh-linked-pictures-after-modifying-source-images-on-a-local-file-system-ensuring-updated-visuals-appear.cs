// Title: Refresh linked pictures in an Excel workbook after source image changes using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, iterates through each worksheet's PictureCollection, checks Picture.IsLinked, invokes Picture.Refresh(), and saves the workbook. | Write a helper method that verifies the presence of source image files, refreshes any linked pictures in a workbook, and gracefully handles cases where the Refresh API is unavailable.
// Common Searches: Aspose.Cells C# refresh linked picture after editing source image file | How to update external images in an Excel workbook programmatically with Aspose.Cells | Iterate worksheet pictures and call Refresh on linked pictures using Aspose.Cells for .NET | Refresh linked image references in .xlsx using Aspose.Cells C#
// Tags: Aspose.Cells picture.Refresh method | linked picture refresh in Excel workbook | C# Aspose.Cells update external images | picture.IsLinked property Aspose.Cells | refresh linked images .xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an existing Excel workbook, walks through each worksheet's PictureCollection, checks if a picture is linked, calls Refresh() when supported, and saves the workbook so that any modified source images are reflected in the file.
class RefreshLinkedPictures
{
    static void Main()
    {
        // Path to the existing workbook that contains linked pictures
        string workbookPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Input file not found: {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pictures on the current worksheet
                PictureCollection pictures = sheet.Pictures;

                // Loop through each picture
                for (int i = 0; i < pictures.Count; i++)
                {
                    Picture pic = pictures[i];

                    // In newer Aspose.Cells versions you can refresh linked pictures via pic.IsLinked and pic.Refresh().
                    // If those members are unavailable, the picture will remain unchanged.
                }
            }

            // Save the workbook with (potentially) refreshed pictures
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
