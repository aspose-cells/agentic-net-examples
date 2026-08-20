// Title: C# – Insert a Linked Picture, Update Its Source URL, and Refresh with Aspose.Cells
// Description: Demonstrates how to add a linked picture from a web URL to a worksheet, modify the picture's SourceFullName to a new URL, trigger a refresh, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells linked picture | C# update picture source | refresh external image Excel | change SourceFullName Aspose | programmatic image replacement
// Common Searches: how to change linked picture URL in Aspose.Cells C# | refresh external image after source change Excel | add linked picture and update its source programmatically | replace web image in generated workbook
// Developer Intent: Replace the URL of an existing linked image in a worksheet and ensure the new picture appears when the file is opened.
// Use Cases: Generate a report with a placeholder image and swap it for a final graphic at runtime. | Automate bulk updates of external logos across many Excel files without recreating shapes. | Create dynamic dashboards where chart images are refreshed based on latest web resources.
// AI Prompts: Write C# code with Aspose.Cells that inserts a linked picture, changes its SourceFullName to a different URL, and saves the workbook. | Explain why calling Workbook.CalculateFormula() helps refresh a linked picture after its source is updated. | Suggest robust error‑handling patterns for updating linked picture URLs in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a linked picture from a web URL to a worksheet, modify the picture's SourceFullName to a new URL, trigger a refresh, and save the workbook using Aspose.Cells for .NET.
class LinkedPictureRefreshDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Initial image URL (linked picture)
            string initialImageUrl = "https://example.com/initial-image.jpg";

            // Insert the linked picture at row 1, column 1 with size 150x150 pixels
            Picture linkedPicture = worksheet.Shapes.AddLinkedPicture(1, 1, 150, 150, initialImageUrl);

            // Verify that the picture is linked
            Console.WriteLine("Is linked picture: " + linkedPicture.IsLink);
            Console.WriteLine("Current SourceFullName: " + linkedPicture.SourceFullName);

            // Change the source URL to a new image
            string newImageUrl = "https://example.com/updated-image.png";
            linkedPicture.SourceFullName = newImageUrl;

            // Force formula recalculation (harmless for picture refresh)
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "LinkedPictureUpdated.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved with updated linked picture at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
