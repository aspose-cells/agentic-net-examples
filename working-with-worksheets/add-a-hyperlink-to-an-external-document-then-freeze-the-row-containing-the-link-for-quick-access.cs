// Title: Add an External PDF Hyperlink and Freeze Its Row with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a hyperlink to an external PDF in cell A2, sets display text and screen tip, freezes the first two rows so the link stays visible, and saves the file as HyperlinkAndFreezeDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# hyperlink | external PDF link | FreezePanes | freeze rows | Excel hyperlink Aspose | worksheet freeze row | add hyperlink Aspose.Cells | set hyperlink text | hyperlink screen tip
// Common Searches: Aspose.Cells add hyperlink to external file | C# FreezePanes example Aspose.Cells | How to freeze row with hyperlink in Excel using Aspose | Set hyperlink display text and screen tip Aspose.Cells | Create clickable PDF link in Excel with Aspose.Cells
// Developer Intent: Generate an Excel workbook that contains a clickable link to an external PDF and keeps the link row pinned by freezing it.
// Use Cases: Report templates with a persistent link to a user‑guide PDF. | Dashboard sheets where the top row provides quick access to supporting documents. | Exported data files that reference external specifications while keeping the reference visible during scrolling.
// AI Prompts: Write C# code with Aspose.Cells to add a hyperlink to a Word document in cell B1, set custom display text and a screen tip, and freeze the top three rows. | Explain the parameters of the FreezePanes method in Aspose.Cells and how to use a cell reference to freeze specific rows. | Provide a step‑by‑step tutorial for adding multiple external hyperlinks to a worksheet and freezing the header row using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace HyperlinkAndFreezeDemo
{
    // Creates a new workbook, inserts a hyperlink to an external PDF in cell A2, sets display text and screen tip, freezes the first two rows so the link stays visible, and saves the file as HyperlinkAndFreezeDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell where the hyperlink will be placed (e.g., A2)
            string hyperlinkCell = "A2";

            // Add a hyperlink that points to an external document (e.g., a PDF file)
            // Parameters: start cell, rows in range, columns in range, address
            worksheet.Hyperlinks.Add(hyperlinkCell, 1, 1, @"C:\Docs\ExternalDocument.pdf");

            // Optionally set display text and screen tip for better UX
            int linkIndex = worksheet.Hyperlinks.Count - 1; // index of the hyperlink just added
            Hyperlink link = worksheet.Hyperlinks[linkIndex];
            link.TextToDisplay = "Open External Document";
            link.ScreenTip = "Click to open the PDF file";

            // Freeze the row that contains the hyperlink (row 2)
            // FreezePanes(string cellName, int freezedRows, int freezedColumns)
            // Use cell "A3" as the split point and freeze the first 2 rows
            worksheet.FreezePanes("A3", 2, 0);

            // Save the workbook to an Excel file
            workbook.Save("HyperlinkAndFreezeDemo.xlsx");
        }
    }
}
