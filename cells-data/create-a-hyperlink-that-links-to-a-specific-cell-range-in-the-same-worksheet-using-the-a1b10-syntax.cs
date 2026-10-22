// Title: Create an internal hyperlink to the A1:B10 range in the same worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a hyperlink in cell C1 which navigates to the A1:B10 range on the same worksheet. | Show how to set a screen tip for an internal worksheet hyperlink using Aspose.Cells. | Demonstrate creating a workbook, inserting a '#A1:B10' internal hyperlink, and saving it as HyperlinkToRange.xlsx.
// Common Searches: Aspose.Cells C# add hyperlink that jumps to a specific range in the same sheet | how to use '#A1:B10' syntax for internal hyperlinks with Aspose.Cells | setting screen tip for an internal hyperlink in Aspose.Cells .NET | example of creating a hyperlink to a cell range using Aspose.Cells for C#
// Tags: Aspose.Cells add hyperlink to worksheet range | Aspose.Cells set hyperlink screen tip | Aspose.Cells '#A1:B10' address syntax | Aspose.Cells internal hyperlink example

using Aspose.Cells;
using System;

// Creates a new workbook, inserts a hyperlink in cell C1 that points to the A1:B10 range on the same worksheet using the '#A1:B10' syntax, assigns a screen tip, and saves the file as HyperlinkToRange.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell where the hyperlink will be placed (e.g., C1)
        int hyperlinkRow = 0;      // Row index for C1 (0‑based)
        int hyperlinkColumn = 2;   // Column index for C1 (0‑based)

        // Add a hyperlink on cell C1 that points to the range A1:B10 in the same sheet
        // The address uses the '#A1:B10' syntax for internal references
        int hyperlinkIndex = sheet.Hyperlinks.Add(
            hyperlinkRow,      // first row of the hyperlink range (C1)
            hyperlinkColumn,   // first column of the hyperlink range (C1)
            1,                 // number of rows the hyperlink occupies
            1,                 // number of columns the hyperlink occupies
            "#A1:B10");        // target range within the same worksheet

        // Optional: set a screen tip for the hyperlink
        sheet.Hyperlinks[hyperlinkIndex].ScreenTip = "Jump to range A1:B10";

        // Save the workbook to a file
        workbook.Save("HyperlinkToRange.xlsx");
    }
}
