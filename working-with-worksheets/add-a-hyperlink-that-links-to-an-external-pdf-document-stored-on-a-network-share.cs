// Title: Add a UNC‑Path PDF hyperlink to an Excel cell using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to create a new Workbook with Aspose.Cells, insert a hyperlink in cell A1 that points to a PDF on a Windows network share (UNC path), set custom display text and screen tip, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells C# hyperlink UNC | Excel link to network share PDF | Aspose.Cells add external PDF link | set hyperlink display text Aspose.Cells | save workbook with hyperlink | Windows file server Excel link
// Common Searches: Aspose.Cells add hyperlink to UNC path | C# create Excel link to network share PDF | How to set screen tip for Aspose.Cells hyperlink | Save Excel workbook after adding external link Aspose.Cells | Open PDF from Excel using network share path
// Developer Intent: Insert an Excel cell hyperlink that opens a PDF located on a corporate file server via a UNC path.
// Use Cases: Provide quick access to policy PDFs stored on a shared drive from generated financial reports. | Build a dashboard that lists project documentation, each row linking to its PDF on a central server. | Create a template workbook for field technicians that includes clickable manuals hosted on the company’s network share. | Automate batch insertion of multiple UNC‑based PDF links into a summary worksheet.
// AI Prompts: Generate C# Aspose.Cells code to add a hyperlink to cell B2 that points to \\Server\Docs\Report.pdf, with display text 'View Report' and a screen tip. | Explain how to configure Aspose.Cells to use a UNC path for an external PDF and ensure the link works when the workbook is opened on a different client. | Write a loop in C# that reads a list of network‑share PDF paths from a CSV and adds each as a hyperlink to successive rows in an Aspose.Cells worksheet. | Show how to set hyperlink properties (TextToDisplay, ScreenTip) and then save the workbook in .xlsm format while preserving macros.

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExample
{
    // This C# example shows how to create a new Workbook with Aspose.Cells, insert a hyperlink in cell A1 that points to a PDF on a Windows network share (UNC path), set custom display text and screen tip, and save the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell where the hyperlink will be placed
            string cellName = "A1";

            // Define the UNC path to the external PDF on a network share
            // Example: \\ServerName\SharedFolder\Document.pdf
            string pdfPath = @"\\ServerName\SharedFolder\Document.pdf";

            // Add a hyperlink to the specified cell.
            // Parameters: cell name, rows in range, columns in range, address (UNC path)
            int hyperlinkIndex = worksheet.Hyperlinks.Add(cellName, 1, 1, pdfPath);

            // Retrieve the created hyperlink to customize its display text and screen tip
            Hyperlink hyperlink = worksheet.Hyperlinks[hyperlinkIndex];
            hyperlink.TextToDisplay = "Open PDF Document";
            hyperlink.ScreenTip = "Click to open the PDF located on the network share";

            // Optionally put a placeholder value in the cell (the hyperlink text will be shown)
            worksheet.Cells[cellName].PutValue(hyperlink.TextToDisplay);

            // Save the workbook to an Excel file
            workbook.Save("WorkbookWithNetworkPdfLink.xlsx");
        }
    }
}
