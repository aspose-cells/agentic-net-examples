// Title: Aspose.Cells for .NET – Load workbook from stream, insert linked picture at G10, refresh external links
// Description: Shows how to open an Excel file from a .NET stream, place a web‑linked image in cell G10, update all external data connections, and save the workbook back to a stream.
// Keywords: Aspose.Cells | C# | load workbook from stream | add linked picture | linked image Excel | cell G10 | refresh external links | UpdateLinkedDataSource | SaveFormat.Xlsx | web image Excel
// Common Searches: Aspose.Cells add linked picture to a specific cell | Refresh external links after inserting an image with Aspose.Cells | Load Excel from memory stream and embed web image C# | How to use AddLinkedPicture in Aspose.Cells | UpdateLinkedDataSource example Aspose.Cells .NET
// Developer Intent: Open a workbook from a stream, embed a URL‑based picture at G10, refresh any linked data sources, and write the result to an output stream.
// Use Cases: Import a template workbook from a memory stream, attach a product thumbnail from a CDN to cell G10, refresh data links, and generate a final report. | Process user‑uploaded Excel files, add a corporate logo hosted online to the header cell G10, and ensure all external connections are up‑to‑date before saving. | Create a dynamic sales dashboard where each row receives a linked image from a web service, requiring a link refresh after insertion.
// AI Prompts: Generate C# code that loads an Excel workbook from a stream, adds a linked picture from a URL to cell G10, calls UpdateLinkedDataSource to refresh external links, and saves the file to an output stream. | Explain step‑by‑step how Aspose.Cells for .NET can insert a web‑linked image into a specific cell and then refresh all external data connections.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to open an Excel file from a .NET stream, place a web‑linked image in cell G10, update all external data connections, and save the workbook back to a stream.
class Program
{
    static void Main()
    {
        // Load the workbook from an input stream (replace with your actual stream source)
        using (Stream inputStream = File.OpenRead("input.xlsx"))
        {
            Workbook workbook = new Workbook(inputStream);

            // Access the first worksheet (or any specific worksheet as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the image URL to be linked
            string imageUrl = "https://example.com/sample.jpg";

            // Cell G10 corresponds to row index 9 and column index 6 (zero‑based)
            int topRow = 9;      // Row 10 in Excel
            int leftColumn = 6;  // Column G in Excel
            int height = 100;    // Height in pixels
            int width = 100;     // Width in pixels

            // Add a linked picture to the worksheet at the specified cell
            worksheet.Shapes.AddLinkedPicture(topRow, leftColumn, height, width, imageUrl);

            // Refresh any external links present in the workbook
            workbook.UpdateLinkedDataSource(new Workbook[0]);

            // Save the modified workbook to an output stream (replace with your actual destination)
            using (Stream outputStream = File.Create("output.xlsx"))
            {
                workbook.Save(outputStream, SaveFormat.Xlsx);
            }
        }
    }
}
