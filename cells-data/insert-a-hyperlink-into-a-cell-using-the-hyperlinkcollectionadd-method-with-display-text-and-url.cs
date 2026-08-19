// Title: Insert a Hyperlink with Custom Text and Screen Tip into an Excel Cell using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, access the first worksheet, and add a hyperlink to cell A1 with a URL, display text, and a screen tip using the HyperlinkCollection.Add method, then save the file as HyperlinkWithDisplayText.xlsx.
// Keywords: Aspose.Cells | C# | .NET | add hyperlink | custom display text | screen tip | HyperlinkCollection.Add | Excel automation | programmatic hyperlink
// Common Searches: Aspose.Cells add hyperlink with display text C# | How to set screen tip for Excel hyperlink using Aspose | HyperlinkCollection.Add parameters example | Insert clickable link in Excel via .NET code | Create hyperlink with tooltip in Aspose.Cells
// Developer Intent: Add a hyperlink that shows custom text and a tooltip to a specific Excel cell programmatically.
// Use Cases: Generate a report where a cell contains a labeled link that opens a website and displays a helpful tooltip. | Build navigation links between worksheets with descriptive text for better user experience. | Create Excel templates that guide users with clickable labels and explanatory screen tips.
// AI Prompts: Show how to add multiple hyperlinks with different display texts and screen tips to a range of cells using Aspose.Cells for .NET. | Explain how to update the URL or screen tip of an existing hyperlink in an Excel workbook with C#. | Provide code to read all hyperlinks from a worksheet and list their addresses, display texts, and screen tips.

using Aspose.Cells;

// Demonstrates how to create a workbook, access the first worksheet, and add a hyperlink to cell A1 with a URL, display text, and a screen tip using the HyperlinkCollection.Add method, then save the file as HyperlinkWithDisplayText.xlsx.
class HyperlinkExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell A1 with display text and a screen tip
        // Parameters: start cell, end cell, URL, text to display, screen tip
        worksheet.Hyperlinks.Add("A1", "A1", "https://www.aspose.com", "Visit Aspose", "Open Aspose website");

        // Save the workbook to a file
        workbook.Save("HyperlinkWithDisplayText.xlsx");
    }
}
