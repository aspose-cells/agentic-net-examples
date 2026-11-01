// Title: Add an external hyperlink to cell H2 in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new workbook, writes "Example Link" to cell H2, adds a hyperlink to https://example.com, and saves the file. | Write a program using Aspose.Cells to set custom display text and an external URL hyperlink for cell H2, then export as HyperlinkExample.xlsx.
// Common Searches: Aspose.Cells C# add hyperlink to cell H2 with custom text | how to insert external URL hyperlink into Excel using Aspose.Cells .NET | C# Aspose.Cells create workbook and set hyperlink for specific cell
// Tags: Aspose.Cells add hyperlink to Excel cell | C# set external URL hyperlink in worksheet | hyperlink cell H2 Aspose.Cells example | save workbook with hyperlink using Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a new workbook, writes "Example Link" into cell H2, attaches an external hyperlink to https://example.com, and saves the file as HyperlinkExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell coordinates for H2 (zero‑based indices)
        int row = 1;      // Row 2
        int column = 7;   // Column H

        // Optional: set display text for the hyperlink
        sheet.Cells[row, column].PutValue("Example Link");

        // Add a hyperlink to the external URL
        sheet.Hyperlinks.Add(row, column, 1, 1, "https://example.com");

        // Save the workbook
        workbook.Save("HyperlinkExample.xlsx");
    }
}
