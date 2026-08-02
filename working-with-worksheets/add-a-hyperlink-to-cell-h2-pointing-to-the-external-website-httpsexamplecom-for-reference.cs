// Title: Add an External Hyperlink to Cell H2 with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, access the first worksheet, insert a hyperlink to https://example.com in cell H2, set the display text to "Example Site", and save the file as HyperlinkInH2.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | hyperlink | cell H2 | external URL | add hyperlink Aspose.Cells | worksheet hyperlink | save workbook | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add hyperlink to cell | C# insert external link in Excel cell using Aspose | How to set hyperlink text in Aspose.Cells | Create hyperlink in H2 with Aspose.Cells .NET | Aspose.Cells Hyperlinks.Add example
// Developer Intent: Insert an external URL hyperlink into cell H2 of an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Embed a reference link in automatically generated financial reports. | Provide quick access to online documentation from exported data sheets. | Create a clickable table‑of‑contents entry that opens a website. | Add marketing or support URLs to cells in bulk‑generated Excel files.
// AI Prompts: Write C# code using Aspose.Cells to add a hyperlink with custom display text to cell H2. | Show how to add multiple external hyperlinks to different cells in a worksheet with Aspose.Cells for .NET. | Explain how to modify the address and display text of an existing hyperlink in an Aspose.Cells workbook. | Provide a step‑by‑step guide to create, style, and validate hyperlinks in Excel using Aspose.Cells C# API.

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExample
{
    // Demonstrates how to create a new workbook, access the first worksheet, insert a hyperlink to https://example.com in cell H2, set the display text to "Example Site", and save the file as HyperlinkInH2.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a hyperlink to cell H2 (row 1, column 7) pointing to the external website
            // Parameters: cell name, number of rows, number of columns, hyperlink address
            worksheet.Hyperlinks.Add("H2", 1, 1, "https://example.com");

            // Optionally set the display text for the hyperlink
            worksheet.Hyperlinks[0].TextToDisplay = "Example Site";

            // Save the workbook to a file
            workbook.Save("HyperlinkInH2.xlsx");
        }
    }
}
