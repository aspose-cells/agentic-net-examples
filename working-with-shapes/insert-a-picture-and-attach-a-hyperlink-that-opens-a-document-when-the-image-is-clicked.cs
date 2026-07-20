// Title: C# – Insert a Linked Picture with Clickable Hyperlink in Excel using Aspose.Cells
// Description: Creates a new workbook, adds a linked image to a specific cell, attaches a file‑URI hyperlink that opens a PDF when the picture is clicked, and saves the file as .xlsx.
// Keywords: Aspose.Cells C# | AddLinkedPicture | picture hyperlink Excel | Excel image link programmatically | C# Aspose.Cells example | linked picture Excel | file URI hyperlink | AddHyperlink Aspose.Cells | insert image with click action | Excel dashboard navigation
// Common Searches: Aspose.Cells add linked picture with hyperlink C# | Insert image that opens PDF when clicked in Excel using Aspose.Cells | C# code to attach hyperlink to a picture in Excel | How to use AddLinkedPicture and AddHyperlink in Aspose.Cells | Create clickable image button in Excel via Aspose.Cells .NET
// Developer Intent: Add an image to a worksheet and make it open a document on click.
// Use Cases: Interactive dashboard icons that launch detailed PDF reports. | Product thumbnails that open specification sheets when selected. | Navigation sheet where pictures act as buttons to open related files.
// AI Prompts: Generate C# code with Aspose.Cells to insert a linked picture and assign a file‑URI hyperlink to a PDF. | Show how to calculate picture size and position based on cell dimensions when using AddLinkedPicture. | Provide best‑practice error handling for missing image or target document paths in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a linked image to a specific cell, attaches a file‑URI hyperlink that opens a PDF when the picture is clicked, and saves the file as .xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Path to the image file that will be linked
        string imagePath = @"C:\Images\sample.jpg";

        // Add a linked picture to the worksheet at row 2, column 2 (zero‑based indices)
        // Height and width are specified in pixels
        Picture picture = worksheet.Shapes.AddLinkedPicture(
            topRow: 1,          // row index (0‑based)
            leftColumn: 1,      // column index (0‑based)
            height: 100,        // picture height in pixels
            width: 100,         // picture width in pixels
            sourceFullName: imagePath);

        // Hyperlink that opens a document when the picture is clicked
        // Use a file URI or any valid URL
        string documentUri = "file:///C:/Docs/TargetDocument.pdf";
        picture.AddHyperlink(documentUri);

        // Save the workbook to a file
        workbook.Save("Workbook_With_Picture_Hyperlink.xlsx");
    }
}
