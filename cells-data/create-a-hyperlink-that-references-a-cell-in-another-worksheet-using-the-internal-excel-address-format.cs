// Title: Add an internal worksheet hyperlink to a cell with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, add a second worksheet, write a value to Sheet2!A1, insert a hyperlink in Sheet1!B2 that points to Sheet2!A1 using the internal Excel address format, set the hyperlink's display text, and save the file.
// Keywords: Aspose.Cells | C# hyperlink worksheet | internal Excel address format | link to another sheet | Hyperlink.TextToDisplay | Aspose.Cells example | Excel navigation hyperlink
// Common Searches: Aspose.Cells add hyperlink to another worksheet | C# internal address format hyperlink Excel | how to link Sheet1 cell to Sheet2 cell using Aspose.Cells | set display text for worksheet hyperlink Aspose.Cells | create table of contents with hyperlinks Aspose.Cells
// Developer Intent: Insert a hyperlink in one worksheet that navigates to a specific cell on a different worksheet using Aspose.Cells for .NET.
// Use Cases: Build a table‑of‑contents sheet that links to sections across multiple worksheets. | Enable quick navigation from a dashboard to detailed data sheets. | Create interactive reports where clicking a cell opens a related chart or table on another sheet.
// AI Prompts: Generate C# code with Aspose.Cells to add a hyperlink from Sheet1!C5 to Sheet3!D10 and set the display text to "Open Details". | Explain the internal address format for worksheet hyperlinks in Aspose.Cells and how to reference named ranges. | Write code that iterates over a list of worksheet names and adds hyperlinks to each sheet's A1 cell from a master index sheet.

using System;
using Aspose.Cells;

// Demonstrates how to create a new workbook, add a second worksheet, write a value to Sheet2!A1, insert a hyperlink in Sheet1!B2 that points to Sheet2!A1 using the internal Excel address format, set the hyperlink's display text, and save the file.
class HyperlinkToAnotherSheet
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a second worksheet (default name will be "Sheet2")
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet1 = workbook.Worksheets[0];
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];

        // Put a value in the target cell on Sheet2
        sheet2.Cells["A1"].PutValue("Target Cell");

        // Add a hyperlink in Sheet1!B2 that points to Sheet2!A1 using internal address format
        sheet1.Hyperlinks.Add("B2", 1, 1, "Sheet2!A1");

        // Optionally set the display text of the hyperlink
        Hyperlink hyperlink = sheet1.Hyperlinks[0];
        hyperlink.TextToDisplay = "Go to Sheet2 A1";

        // Save the workbook
        workbook.Save("HyperlinkToOtherSheet.xlsx");
    }
}
