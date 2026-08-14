// Title: Create a UNC‑Path PDF Hyperlink in an Excel Cell with Aspose.Cells for .NET (C#)
// Description: This C# example uses Aspose.Cells to generate a new workbook, adds a hyperlink in cell A1 that points to a PDF on a network share (\\Server\Share\Document.pdf), sets the display text to "Open PDF", and saves the file as NetworkPdfHyperlink.xlsx. [PDF Link](file://\\Server\Share\Document.pdf)
// Keywords: Aspose.Cells hyperlink UNC | C# Excel network share link | add PDF hyperlink Aspose.Cells | Worksheet.Hyperlinks.Add example | .NET Excel external document link | Excel cell hyperlink to file server
// Common Searches: Aspose.Cells add hyperlink to network PDF | C# Excel UNC path hyperlink example | How to link a cell to a file server document | Set hyperlink display text in Aspose.Cells | Save Excel with external PDF link using .NET
// Developer Intent: Insert a hyperlink in an Excel worksheet cell that opens a PDF located on a corporate network share.
// Use Cases: Embed quick access to policy documents stored on a file server within generated reports. | Create a catalog where each row links to its corresponding specification PDF on a shared drive. | Distribute Excel dashboards that reference external manuals or guidelines hosted on a network location.
// AI Prompts: Write C# code with Aspose.Cells to add a UNC‑path hyperlink to a PDF and customize the link text. | Explain security considerations when using Worksheet.Hyperlinks.Add with network shares. | Show how to add multiple cells each linking to different PDFs on a shared server using a loop.

using System;
using Aspose.Cells;

// This C# example uses Aspose.Cells to generate a new workbook, adds a hyperlink in cell A1 that points to a PDF on a network share (\\Server\Share\Document.pdf), sets the display text to "Open PDF", and saves the file as NetworkPdfHyperlink.xlsx. [PDF Link](file://\\Server\Share\Document.pdf)
class AddNetworkPdfHyperlink
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the cell that will contain the hyperlink
        string cellName = "A1";

        // UNC path to the external PDF on a network share
        string pdfAddress = @"\\Server\Share\Document.pdf";

        // Add the hyperlink to the specified cell (A1)
        // Parameters: cell name, rows in range, columns in range, address
        worksheet.Hyperlinks.Add(cellName, 1, 1, pdfAddress);

        // Set the display text for the hyperlink
        worksheet.Cells[cellName].PutValue("Open PDF");
        Hyperlink hyperlink = worksheet.Hyperlinks[0];
        hyperlink.TextToDisplay = "Open PDF";

        // Save the workbook
        workbook.Save("NetworkPdfHyperlink.xlsx");
    }
}
