// Title: Validate target="_parent" in Aspose.Cells HTML output with HtmlSaveOptions.LinkTargetType=Parent (C#)
// Description: This C# example creates a workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Parent (generating target="_parent"), saves the file as HTML, reads the output, and confirms the hyperlink contains the expected target attribute.
// Keywords: Aspose.Cells | HtmlSaveOptions | LinkTargetType | Parent target | target=_parent | C# HTML export | hyperlink validation | frame navigation | generated HTML check
// Common Searches: Aspose.Cells set hyperlink target to _parent | validate target attribute in saved HTML | HtmlSaveOptions LinkTargetType Parent example | C# check generated HTML for target=_parent | Aspose.Cells HTML export link behavior
// Developer Intent: Confirm that the HTML file produced by Aspose.Cells includes a hyperlink with target="_parent" when the LinkTargetType option is set to Parent.
// Use Cases: Generating HTML reports where links must open in the parent frame of a frameset. | Automated regression tests that verify correct link target rendering after HTML export. | Configuring link behavior for web pages embedding Aspose.Cells output inside iframes or frames.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as HTML with LinkTargetType set to Parent and asserts the presence of target="_parent" in the output. | Create an NUnit test that loads the saved HTML from Aspose.Cells and checks that all hyperlinks contain target="_parent" when HtmlSaveOptions.LinkTargetType is Parent. | Explain how HtmlSaveOptions.LinkTargetType influences hyperlink targets in the exported HTML and provide a snippet that verifies the attribute.

using System;
using System.IO;
using Aspose.Cells;

// This C# example creates a workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Parent (generating target="_parent"), saves the file as HTML, reads the output, and confirms the hyperlink contains the expected target attribute.
class HtmlLinkTargetValidation
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put some text in a cell and add a hyperlink to it
        sheet.Cells["A1"].PutValue("Visit Aspose");
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Configure HTML save options to use the Parent target type
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Parent; // target="_parent"

        // Define the output HTML file path
        string htmlPath = "LinkTargetParent.html";

        // Save the workbook as HTML
        workbook.Save(htmlPath, saveOptions);

        // Read the generated HTML content
        string htmlContent = File.ReadAllText(htmlPath);

        // Verify that the hyperlink contains target="_parent"
        bool containsParentTarget = htmlContent.Contains("target=\"_parent\"");

        // Output the validation result
        Console.WriteLine(containsParentTarget
            ? "Validation succeeded: target=\"_parent\" is present."
            : "Validation failed: target=\"_parent\" is missing.");
    }
}
