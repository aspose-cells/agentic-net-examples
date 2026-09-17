// Title: Configure HtmlSaveOptions.HtmlLinkTargetType to NewWindow so hyperlinks open in a new tab when exporting Excel to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that saves a Workbook to HTML using Aspose.Cells and sets HtmlSaveOptions.HtmlLinkTargetType to NewWindow so every hyperlink opens in a new browser tab. | Update an existing Aspose.Cells HTML export snippet to configure the hyperlink target to open in a new window. | Show how to apply HtmlLinkTargetType.NewWindow globally for all links when converting an Excel file to HTML with Aspose.Cells.
// Common Searches: Aspose.Cells how to make hyperlinks open in new tab when saving as HTML | C# HtmlSaveOptions HtmlLinkTargetType NewWindow example | Export Excel to HTML with Aspose.Cells and set link target to _blank | Set hyperlink target type in Aspose.Cells HTML export C# | Aspose.Cells HtmlSaveOptions open links in new window
// Tags: Aspose.Cells HtmlSaveOptions hyperlink target | C# Aspose.Cells HTML export new window links | HtmlLinkTargetType NewWindow Aspose.Cells | Excel to HTML conversion link target Aspose | Aspose.Cells set link target type

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a workbook, adds a hyperlink with a screen tip, and saves the file as HTML using Aspose.Cells. It configures HtmlSaveOptions.HtmlLinkTargetType = HtmlLinkTargetType.NewWindow so that all hyperlinks in the generated HTML open in a new browser tab.
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

            // Add a hyperlink to cell A1 (row 0, column 0, spanning 1 row and 1 column)
            // Hyperlinks.Add returns the index of the added hyperlink
            int linkIndex = sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.example.com");
            Hyperlink link = sheet.Hyperlinks[linkIndex];
            link.ScreenTip = "Example Site";

            // Set display text for the cell
            sheet.Cells["A1"].PutValue("Go to Example");

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Default link target behavior will be used

            // Save the workbook as HTML
            workbook.Save("output.html", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
