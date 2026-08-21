// Title: C# Aspose.Cells: Add an external PDF hyperlink and freeze its row
// Description: Creates a new workbook, inserts a hyperlink to an external PDF in cell A2, customizes the display text and screen tip, freezes row 2 with FreezePanes, and saves the file as HyperlinkAndFreeze.xlsx.
// Keywords: Aspose.Cells add hyperlink | C# Excel external link | FreezePanes row | hyperlink display text Aspose.Cells | freeze row after hyperlink
// Common Searches: Aspose.Cells add hyperlink to external file | Freeze specific row in Aspose.Cells C# | Set screen tip for hyperlink Aspose.Cells | FreezePanes using cell reference Aspose.Cells | C# Excel hyperlink external PDF
// Developer Intent: Insert an external document link into a worksheet cell and keep the link visible by freezing its row.
// Use Cases: Provide quick access to policy PDFs from a summary sheet while keeping the link row fixed. | Create a navigation worksheet where each row links to a separate report and the link rows stay in view. | Build a dashboard that pins rows containing external resource links for constant accessibility.
// AI Prompts: Generate C# Aspose.Cells code that adds a hyperlink to a Word document in cell B5, sets custom display text, and freezes the first three rows. | Show how to add multiple external hyperlinks to different cells and freeze the top header row in an Aspose.Cells workbook. | Explain the use of FreezePanes with a cell reference to lock rows after inserting hyperlinks in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new workbook, inserts a hyperlink to an external PDF in cell A2, customizes the display text and screen tip, freezes row 2 with FreezePanes, and saves the file as HyperlinkAndFreeze.xlsx.
class HyperlinkAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to an external document in cell A2
        // Parameters: cell name, total rows, total columns, address (file path)
        int hyperlinkIndex = worksheet.Hyperlinks.Add("A2", 1, 1, @"C:\Docs\ExternalDoc.pdf");

        // Optionally set the display text and screen tip for the hyperlink
        Hyperlink link = worksheet.Hyperlinks[hyperlinkIndex];
        link.TextToDisplay = "Open External Document";
        link.ScreenTip = "Click to open the external PDF";

        // Freeze the row that contains the hyperlink (row 2) for quick access
        // FreezePanes(cellName, freezedRows, freezedColumns)
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("HyperlinkAndFreeze.xlsx");
    }
}
