// Title: Add a UNC network‑share hyperlink to an Excel cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink to UNC path in Excel worksheet | how to create a network share link in an Excel file with Aspose.Cells | set screen tip for a UNC hyperlink using Aspose.Cells API | save Excel workbook with UNC hyperlink using Aspose.Cells .NET | Aspose.Cells Hyperlinks.Add example for network file path
// Tags: Aspose.Cells hyperlink insertion API | Hyperlinks.Add method parameters | Excel cell link to network location | configure hyperlink tooltip Aspose.Cells | persist workbook with embedded links Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, inserts a hyperlink in cell A1 that points to a UNC network share (\\ServerName\SharedFolder\Document.pdf), assigns a screen tip, sets display text, and saves the file as HyperlinkExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // UNC path to the network share
            string uncPath = @"\\ServerName\SharedFolder\Document.pdf";

            // Add a hyperlink to cell A1 (row 0, column 0)
            // The Add method returns the index of the created hyperlink
            int hyperlinkIndex = sheet.Hyperlinks.Add(0, 0, 1, 1, uncPath);
            Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
            hyperlink.ScreenTip = "Open Document";

            // Set the display text for the hyperlink
            sheet.Cells[0, 0].PutValue("Open Document");

            // Save the workbook to a file
            workbook.Save("HyperlinkExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
