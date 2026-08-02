// Title: Load Workbook from Stream, Insert Linked Picture into Cell G10, and Refresh Links – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an Excel file from a FileStream, place a web‑linked image into cell G10 (row 9, column 6) with a 100 × 100 px size, refresh all external links—including the picture—using UpdateLinkedDataSource, and save the result.
// Keywords: Aspose.Cells load workbook from stream | C# add linked picture Excel | Insert web image into cell G10 | Refresh external links Aspose.Cells | UpdateLinkedDataSource example | Excel linked picture C# | Aspose.Cells shape API
// Common Searches: How to add a linked picture to a specific cell with Aspose.Cells .NET | Refresh external links after inserting an image in Excel using Aspose.Cells | Load Excel from MemoryStream and embed a URL image in C#
// Developer Intent: Load an Excel workbook from a stream, attach a web‑linked picture to cell G10, then update the workbook’s external links.
// Use Cases: Automated sales reports that pull product thumbnails from a CDN into predefined cells and ensure the links stay current. | Marketing dashboards that display live images in cell G10, requiring batch processing with automatic link refresh. | Template generation scripts that embed external graphics and need to synchronize linked data before distribution.
// AI Prompts: Generate C# code with Aspose.Cells to read an Excel file from a MemoryStream, add a linked picture from a URL to cell G10, refresh all external links, and write the workbook to disk. | Explain the purpose of Workbook.UpdateLinkedDataSource in Aspose.Cells and how it affects linked pictures and other external data sources.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to open an Excel file from a FileStream, place a web‑linked image into cell G10 (row 9, column 6) with a 100 × 100 px size, refresh all external links—including the picture—using UpdateLinkedDataSource, and save the result.
class Program
{
    static void Main()
    {
        // Load the workbook from a file stream (replace with any input stream as needed)
        using (FileStream inputStream = new FileStream("input.xlsx", FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(inputStream);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a linked picture to cell G10 (zero‑based row 9, column 6)
            // Height and width are set to 100 pixels each
            string imageUrl = "https://example.com/sample.jpg";
            worksheet.Shapes.AddLinkedPicture(9, 6, 100, 100, imageUrl);

            // Refresh external links (including the linked picture)
            workbook.UpdateLinkedDataSource(new Workbook[0]);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
