// Title: C# – Insert a Linked Picture and Update Its Source URL with Aspose.Cells
// Description: Demonstrates how to add a linked picture to a worksheet, change its SourceFullName to a new image URL, and save the workbook so the picture reflects the updated source.
// Keywords: Aspose.Cells linked picture | C# update picture URL | .NET insert linked image | SourceFullName property | refresh external image | worksheet picture sample | GitHub Aspose.Cells example | programmatic image replacement
// Common Searches: Aspose.Cells change linked picture URL C# | how to update SourceFullName of a picture in Aspose.Cells | replace linked image in Excel workbook using .NET | refresh external picture after URL change Aspose.Cells | sample code for linked picture in Aspose.Cells
// Developer Intent: Insert a linked picture into a worksheet and then programmatically replace its source URL.
// Use Cases: Create a report template with a placeholder image that is later swapped for a user‑specific URL. | Build dashboards that pull external graphics and need to refresh them when the source files are updated. | Generate workbooks that reference online logos or charts, allowing the URLs to be changed without recreating the file.
// AI Prompts: Show C# code using Aspose.Cells to add a linked picture and later modify its SourceFullName property. | Explain how to force a linked picture to refresh after its URL has been changed in an existing workbook. | Provide error‑handling patterns for invalid or unreachable image URLs when updating a linked picture.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a linked picture to a worksheet, change its SourceFullName to a new image URL, and save the workbook so the picture reflects the updated source.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a linked picture (initial source URL)
        string initialUrl = "https://example.com/initial.jpg";
        Picture linkedPicture = worksheet.Shapes.AddLinkedPicture(
            topRow: 1,          // row index where the picture starts
            leftColumn: 1,      // column index where the picture starts
            height: 200,        // height in pixels
            width: 200,         // width in pixels
            sourceFullName: initialUrl);

        // Change the source URL of the linked picture to a new image
        string newUrl = "https://example.com/updated.jpg";
        linkedPicture.SourceFullName = newUrl;

        // Save the workbook (the picture will now reference the new URL)
        workbook.Save("LinkedPictureUpdated.xlsx");
    }
}
