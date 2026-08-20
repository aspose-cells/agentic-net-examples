// Title: Custom Ribbon Button to Update External Link Paths with Aspose.Cells (C#)
// Description: Demonstrates how to add custom Ribbon XML to a workbook, simulate a button click, and programmatically replace SharePoint base URLs in external links using Aspose.Cells for .NET.
// Keywords: Aspose.Cells external links | custom ribbon button C# | update Excel link paths | replace SharePoint URL Aspose | Workbook.RibbonXml example | modify external data source programmatically
// Common Searches: how to change external link URLs in Aspose.Cells | custom ribbon XML for Excel workbook C# | replace SharePoint path in Excel external links | Aspose.Cells update external links programmatically | simulate ribbon button click Aspose.Cells
// Developer Intent: Validate that a custom Ribbon button correctly rewrites external link URLs in an Aspose.Cells workbook.
// Use Cases: Migrate Excel workbooks after moving SharePoint files to an internal server. | Provide end‑users with a UI button that fixes outdated data source links in one click. | Automate pre‑save checks to ensure all external references point to the new location.
// AI Prompts: Write C# code using Aspose.Cells to iterate Workbook.Worksheets.ExternalLinks and replace a given base URL. | Show how to embed custom Ribbon XML in a workbook and trigger link‑path updates from the button action. | Explain how to unit‑test that external link URLs have been updated after simulating the Ribbon button.

using System;
using Aspose.Cells;

namespace CustomRibbonExternalLinkDemo
{
    // Demonstrates how to add custom Ribbon XML to a workbook, simulate a button click, and programmatically replace SharePoint base URLs in external links using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add an external link formula in cell A1
            // Example link points to a SharePoint location
            sheet.Cells["A1"].Formula = "='https://sharepoint.example.com/Projects/[data.xlsx]Sheet1'!A1";

            // Define custom Ribbon XML with a button that would trigger the update
            // (In a real UI the button's onAction would be linked to a macro or add‑in;
            // here we just set the XML to demonstrate the ribbon customization)
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"External Links\">" +
                "        <group id=\"updateGroup\" label=\"Update Links\">" +
                "          <button id=\"updateButton\" label=\"Update Paths\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Set the RibbonXml property (create rule)
            workbook.RibbonXml = ribbonXml;

            // Simulate the button click: update all external link paths
            // Replace the old SharePoint base URL with a new internal path
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];
                string original = link.OriginalDataSource;
                string updated = original.Replace(
                    "https://sharepoint.example.com/Projects/",
                    "/internal/projects/");

                // Apply the updated path (property rule)
                link.OriginalDataSource = updated;
            }

            // Verify the changes by writing them to the console
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                Console.WriteLine("Modified External Link: " + link.OriginalDataSource);
            }

            // Save the workbook (save rule)
            workbook.Save("CustomRibbonExternalLinkDemo.xlsx");
        }
    }
}
