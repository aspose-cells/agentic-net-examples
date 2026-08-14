// Title: Insert a linked picture from a URL into cell R2 with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, access the first worksheet, and use Shapes.AddLinkedPicture to place a web‑hosted image at cell R2 (row 2, column 18) with a custom size, then save the file as an .xlsx document.
// Keywords: Aspose.Cells | C# | AddLinkedPicture | linked picture | insert image from URL | cell R2 | worksheet image positioning | external image linking | dynamic picture update | Excel automation
// Common Searches: Aspose.Cells add linked picture to specific cell | C# insert image from web URL into Excel worksheet | Place picture at cell R2 using Aspose.Cells | How to use Shapes.AddLinkedPicture in .NET | Excel workbook with external image link C#
// Developer Intent: Add a web‑based linked image to cell R2 of a worksheet and save the workbook.
// Use Cases: Embed a company logo hosted online into generated reports without increasing file size. | Create dashboards that automatically refresh charts when the source image URL changes. | Build template‑driven spreadsheets that pull branding or product images from a CDN at predefined cell locations.
// AI Prompts: Show C# code using Aspose.Cells to add a linked picture from a URL into cell R2 with a width and height of 100 px. | Explain the parameters of Shapes.AddLinkedPicture for positioning an external image in a worksheet. | Give an example of linking an online image so the picture updates automatically when the source file changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, access the first worksheet, and use Shapes.AddLinkedPicture to place a web‑hosted image at cell R2 (row 2, column 18) with a custom size, then save the file as an .xlsx document.
class InsertPictureFromUrl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // URL of the image to insert
        string imageUrl = "https://example.com/sample.jpg";

        // Add a linked picture at cell R2 (row index 1, column index 17)
        // Height and width are set to 100 pixels each
        worksheet.Shapes.AddLinkedPicture(1, 17, 100, 100, imageUrl);

        // Save the workbook
        workbook.Save("WorkbookWithLinkedPicture.xlsx");
    }
}
