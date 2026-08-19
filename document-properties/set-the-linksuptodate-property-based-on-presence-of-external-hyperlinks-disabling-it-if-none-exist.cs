// Title: C# – Set LinksUpToDate Document Property in Aspose.Cells Based on External Hyperlinks
// Description: Creates a workbook, optionally adds an external hyperlink, scans all worksheets for hyperlinks, sets the BuiltInDocumentProperties.LinksUpToDate flag to reflect their presence, and saves the file.
// Keywords: Aspose.Cells | LinksUpToDate | C# | document properties | external hyperlink detection | Excel link status
// Common Searches: Aspose.Cells set LinksUpToDate false | check workbook for hyperlinks C# | toggle LinksUpToDate based on hyperlink count | how to disable link update prompt in generated Excel
// Developer Intent: Identify whether any worksheet contains external hyperlinks and assign the appropriate value to the LinksUpToDate built‑in property.
// Use Cases: Suppress Excel's "Update Links" dialog when the generated file has no external links. | Maintain accurate link status after programmatically adding or removing hyperlinks across multiple sheets. | Provide downstream systems with a correct LinksUpToDate flag for automated processing.
// AI Prompts: Generate a C# routine that returns true if an Aspose.Cells workbook contains at least one external hyperlink. | Show code that sets Workbook.BuiltInDocumentProperties.LinksUpToDate based on hyperlink detection across all worksheets. | Explain the purpose of the LinksUpToDate property in Excel and how Aspose.Cells can manage it.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates a workbook, optionally adds an external hyperlink, scans all worksheets for hyperlinks, sets the BuiltInDocumentProperties.LinksUpToDate flag to reflect their presence, and saves the file.
class SetLinksUpToDateDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add an external hyperlink (comment out this line to test the case with no hyperlinks)
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Check if any worksheet contains external hyperlinks
        bool hasExternalHyperlinks = false;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Hyperlinks.Count > 0)
            {
                hasExternalHyperlinks = true;
                break;
            }
        }

        // Set the LinksUpToDate property based on the presence of hyperlinks
        workbook.BuiltInDocumentProperties.LinksUpToDate = hasExternalHyperlinks;

        // Save the workbook
        workbook.Save("LinksUpToDateResult.xlsx");
    }
}
