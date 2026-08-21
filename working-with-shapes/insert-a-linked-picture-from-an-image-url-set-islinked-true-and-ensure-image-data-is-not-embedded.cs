// Title: Add a Linked Picture from a URL in Aspose.Cells for .NET (IsLink = true, no embedding)
// Description: Creates a new workbook, inserts a picture that references an external image URL using AddLinkedPicture, sets IsLink to true so the image data stays external, and saves the file as a lightweight Excel document.
// Keywords: Aspose.Cells linked picture | AddLinkedPicture .NET | IsLink true | external image URL Excel | prevent image embedding | lightweight workbook | dynamic logo Excel
// Common Searches: Aspose.Cells add picture from URL without embedding | How to set IsLink for a picture in Aspose.Cells | Create Excel file with linked images .NET | Save workbook with external images Aspose.Cells
// Developer Intent: Insert an image that points to an online source and keep it as a reference rather than embedding the binary data in the Excel file.
// Use Cases: Generate reports that pull logos from a CDN, keeping file size minimal. | Build templates where end‑users can change image URLs without re‑saving binary data. | Create dashboards that automatically reflect updates to external graphics.
// AI Prompts: Show C# code to add a linked picture from a URL in Aspose.Cells and confirm the picture is not embedded. | Explain how to verify that linked pictures have IsLink = true and Data is null after insertion. | Describe workbook save options that preserve external image links in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a picture that references an external image URL using AddLinkedPicture, sets IsLink to true so the image data stays external, and saves the file as a lightweight Excel document.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // URL of the image to be linked
        string imageUrl = "https://example.com/sample.jpg";

        // Add a linked picture to the worksheet (row, column, height, width, source URL)
        Picture linkedPicture = worksheet.Shapes.AddLinkedPicture(1, 1, 100, 100, imageUrl);

        // Explicitly set the picture as linked (IsLink = true)
        linkedPicture.IsLink = true;

        // At this point the picture data is not embedded; linkedPicture.Data should be null

        // Save the workbook
        workbook.Save("LinkedPicture.xlsx");
    }
}
