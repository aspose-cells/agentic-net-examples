// Title: Add an external PDF hyperlink to cell A2 and freeze the first two rows in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that inserts a hyperlink to a PDF file in cell A2, sets a custom display text, freezes rows 1‑2, and saves the workbook. | Write a .NET example that creates an Excel file, adds an external document link to a worksheet cell, applies FreezePanes to keep the link row visible, and outputs HyperlinkWithFreeze.xlsx.
// Common Searches: Aspose.Cells for .NET add PDF hyperlink to a cell and freeze top rows | C# code to keep a hyperlink row visible by freezing panes in Excel using Aspose.Cells | How to insert an external document link and apply FreezePanes with Aspose.Cells C#
// Tags: add hyperlink to cell Aspose.Cells | freeze panes rows Aspose.Cells | external PDF link Excel C# | hyperlink display text Aspose.Cells | freeze top rows after hyperlink insertion

using System;
using Aspose.Cells;

// The example creates a new workbook, adds a hyperlink in cell A2 that points to an external PDF file, sets the link's display text, freezes the first two rows so the hyperlink remains visible while scrolling, and saves the file as HyperlinkWithFreeze.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell A2 that points to an external document
        // Overload: Add(string cellName, int totalRows, int totalColumns, string address)
        worksheet.Hyperlinks.Add("A2", 1, 1, @"C:\Docs\ExternalDoc.pdf");

        // Set the display text for the hyperlink (optional)
        worksheet.Hyperlinks[0].TextToDisplay = "Open External Document";

        // Freeze rows above row 3 (i.e., rows 1‑2) so the hyperlink row stays visible while scrolling
        // Overload: FreezePanes(string cellName, int freezedRows, int freezedColumns)
        worksheet.FreezePanes("A3", 2, 0);

        // Save the workbook
        workbook.Save("HyperlinkWithFreeze.xlsx");
    }
}
