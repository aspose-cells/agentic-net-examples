// Title: Insert a linked picture with clickable hyperlink in Excel using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a linked picture at a specific cell range, attach a file‑based hyperlink to the image, and save the file as an .xlsx document with Aspose.Cells C# API.
// Keywords: Aspose.Cells C# picture hyperlink | AddLinkedPicture Aspose.Cells | Excel image click opens document | AddHyperlink to shape Aspose | linked picture Excel .NET | clickable image Excel workbook | Aspose.Cells example PDF hyperlink
// Common Searches: Aspose.Cells add picture with hyperlink C# | How to make an Excel image open a PDF using Aspose | Insert linked picture and attach hyperlink in .NET | Clickable image in Excel workbook Aspose.Cells | AddHyperlink to picture shape Aspose.Cells example
// Developer Intent: Place an image in a worksheet and make it open a target document when the user clicks the picture.
// Use Cases: Product catalog where each thumbnail links to its spec sheet PDF. | Interactive dashboard icons that launch related reports or manuals. | Training workbook with screenshots that open detailed guide files.
// AI Prompts: Generate C# code with Aspose.Cells to insert a picture from a URL and assign a web‑page hyperlink. | Show how to calculate picture dimensions based on cell size and set the position dynamically. | Provide robust error handling for missing image files or invalid hyperlink URIs when adding a picture hyperlink.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureHyperlinkDemo
{
    // Demonstrates how to create a workbook, add a linked picture at a specific cell range, attach a file‑based hyperlink to the image, and save the file as an .xlsx document with Aspose.Cells C# API.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define picture location and size (in pixels)
            int topRow = 2;      // Row index (0‑based)
            int leftColumn = 2;  // Column index (0‑based)
            int pictureHeight = 150;
            int pictureWidth = 150;

            // Path to the image file to be linked (can be a local file or a URL)
            string imagePath = @"C:\Images\sample.jpg";

            // Add a linked picture to the worksheet
            Picture picture = worksheet.Shapes.AddLinkedPicture(topRow, leftColumn, pictureHeight, pictureWidth, imagePath);

            // Attach a hyperlink to the picture that opens a document when clicked
            // Example: opening a PDF document located on the local file system
            string documentHyperlink = @"file:///C:/Documents/TargetDocument.pdf";
            picture.AddHyperlink(documentHyperlink);

            // Save the workbook to a file
            workbook.Save("Workbook_With_Picture_Hyperlink.xlsx");

            Console.WriteLine("Workbook saved successfully with a picture hyperlink.");
        }
    }
}
