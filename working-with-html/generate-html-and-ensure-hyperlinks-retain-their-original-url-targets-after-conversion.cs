// Title: C# Aspose.Cells – Export Excel to HTML with unchanged hyperlink URLs (LinkTargetType.Self)
// Description: Demonstrates how to create a workbook, add a hyperlink, and save it as HTML using Aspose.Cells while keeping the original href unchanged and opening links in the same window.
// Keywords: Aspose.Cells C# HTML export | preserve hyperlink URL Aspose | HtmlSaveOptions LinkTargetType Self | Excel to HTML without link modification | C# save workbook as HTML | hyperlink target self Aspose.Cells
// Common Searches: Aspose.Cells keep hyperlink URL when saving to HTML | HtmlSaveOptions LinkTargetType Self example C# | export Excel as HTML with original links Aspose | C# Aspose.Cells hyperlink target self | how to prevent link rewriting in Aspose.Cells HTML output
// Developer Intent: Generate HTML from an Excel workbook while ensuring hyperlinks retain their original URLs and open in the same window.
// Use Cases: Publishing Excel‑based reports on a website where links must point to exact external pages. | Creating email‑ready HTML content from spreadsheets without altering destination URLs. | Building a web viewer that displays Excel data as HTML and requires accurate hyperlink navigation.
// AI Prompts: Provide C# code using Aspose.Cells to export a workbook to HTML and keep all hyperlink URLs unchanged by setting HtmlSaveOptions.LinkTargetType to Self. | Explain how HtmlSaveOptions.LinkTargetType = Self influences the generated HTML anchor tags. | Show how to add multiple hyperlinks with different display texts and ensure each retains its original URL after HTML conversion.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add a hyperlink, and save it as HTML using Aspose.Cells while keeping the original href unchanged and opening links in the same window.
class Program
{
    static void Main()
    {
        // Create a new workbook and obtain the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set display text for the hyperlink cell
        sheet.Cells["A1"].PutValue("Visit Aspose");

        // Add a hyperlink to cell A1 with the original URL
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Configure HTML save options to keep the original link target (open in the same window)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Self; // ensures the href remains unchanged

        // Save the workbook as an HTML file
        workbook.Save("output.html", saveOptions);
    }
}
