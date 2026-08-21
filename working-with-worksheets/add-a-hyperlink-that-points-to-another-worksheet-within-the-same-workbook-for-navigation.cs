// Title: Create an Internal Worksheet Hyperlink with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, rename sheets, write a value, and add a hyperlink in Main!A1 that jumps to Details!B5, set its display text, and save the file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | internal hyperlink | worksheet navigation | HyperlinkCollection.Add | Excel hyperlink C# | navigate between sheets | add hyperlink to cell | Excel workbook navigation | Aspose.Cells example
// Common Searches: Aspose.Cells add hyperlink to another sheet | C# internal hyperlink Aspose.Cells | How to link worksheets in Aspose.Cells | Set hyperlink display text Aspose.Cells C# | Navigate to cell B5 from sheet Main Aspose.Cells
// Developer Intent: Add a cell hyperlink that navigates to a specific cell on a different worksheet within the same workbook using Aspose.Cells for .NET.
// Use Cases: Create a table‑of‑contents sheet where each entry links to detailed sections on separate worksheets. | Provide quick navigation from a dashboard sheet to data or chart sheets in an automated report. | Build interactive Excel workbooks that let users jump directly to key metrics or analysis tables.
// AI Prompts: Generate C# code with Aspose.Cells to add an internal hyperlink from Main!A1 to Details!B5 and set the display text. | Explain how to use HyperlinkCollection.Add and TextToDisplay to create navigation links between worksheets in Aspose.Cells. | Write a loop in C# that adds a hyperlink on a summary sheet for each worksheet name in a workbook using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to build a workbook, rename sheets, write a value, and add a hyperlink in Main!A1 that jumps to Details!B5, set its display text, and save the file using Aspose.Cells for C#.
class InternalHyperlinkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and rename it
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Main";

        // Add a second worksheet and rename it
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "Details";

        // Put some content in the target cell on the second sheet
        sheet2.Cells["B5"].PutValue("Target Cell");

        // Add a hyperlink in cell A1 of the first sheet that points to Details!B5
        // Using HyperlinkCollection.Add(string cellName, int totalRows, int totalColumns, string address)
        int hyperlinkIndex = sheet1.Hyperlinks.Add("A1", 1, 1, "Details!B5");

        // Set the display text for the hyperlink (optional)
        sheet1.Hyperlinks[hyperlinkIndex].TextToDisplay = "Go to Details";

        // Save the workbook
        workbook.Save("InternalHyperlinkDemo.xlsx");
    }
}
