// Title: Insert a linked picture from a URL into cell R2 using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a web‑based image to an Excel worksheet with Aspose.Cells, place it at cell R2, set its size, and save the file as XLSX.
// Keywords: Aspose.Cells AddLinkedPicture | C# insert image from URL | place picture in Excel cell | linked picture Aspose.Cells | cell R2 image insertion
// Common Searches: Aspose.Cells add picture from web URL | C# insert image into specific Excel cell | How to use AddLinkedPicture in Aspose.Cells | position picture at cell R2 programmatically | link external image in Excel with Aspose
// Developer Intent: Add a linked image hosted at a URL to cell R2 of a worksheet and save the workbook.
// Use Cases: Create product catalogs where each item’s photo is linked from a CDN and anchored to a designated cell. | Generate marketing reports that pull the latest logo from a remote server and place it in the header row. | Automate invoice generation by embedding a supplier’s online logo into a fixed cell for consistent branding.
// AI Prompts: Write C# code with Aspose.Cells to insert a linked picture from a given URL into cell R2, using a width of 120 px and height of 80 px. | Explain how to adjust alignment and scaling of a picture added with AddLinkedPicture in Aspose.Cells. | Provide sample error handling for unreachable or invalid image URLs when using AddLinkedPicture in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a web‑based image to an Excel worksheet with Aspose.Cells, place it at cell R2, set its size, and save the file as XLSX.
class InsertPictureFromUrl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // URL of the image to be linked
        string imageUrl = "https://example.com/sample.jpg";

        // Position the picture at cell R2 (zero‑based indices: row 1, column 17)
        int topRow = 1;      // Row 2
        int leftColumn = 17; // Column R
        int height = 100;    // Height in pixels
        int width = 100;     // Width in pixels

        // Add a linked picture using the ShapeCollection.AddLinkedPicture method
        worksheet.Shapes.AddLinkedPicture(topRow, leftColumn, height, width, imageUrl);

        // Save the workbook
        workbook.Save("output_with_linked_picture.xlsx");
    }
}
