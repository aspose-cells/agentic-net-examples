// Title: Aspose.Cells for .NET – Update External Link URLs & Add a Custom Ribbon Tab (C#)
// Description: C# example that creates a workbook, replaces old SharePoint URLs in all ExternalLinks, injects custom Ribbon XML to add a new tab, and saves the file as a macro‑enabled workbook using Aspose.Cells.
// Keywords: Aspose.Cells external link path | modify external links C# | RibbonXml Aspose.Cells | custom Excel ribbon .NET | sharepoint url replace Aspose | save workbook as xlsm | Aspose.Cells API example
// Common Searches: how to change external link URL in Aspose.Cells | add custom ribbon tab with Aspose.Cells for .NET | replace SharePoint base address in Excel workbook programmatically | set RibbonXml property in C# Aspose.Cells | update workbook external links before saving
// Developer Intent: Programmatically rewrite external link paths and embed a custom Ribbon UI before persisting the workbook.
// Use Cases: Migrate Excel files to a new SharePoint site by updating every external data source URL. | Introduce a company‑specific ribbon tab with custom commands without using VBA. | Generate macro‑enabled (.xlsm) reports that retain UI customizations across deployments.
// AI Prompts: Write C# code with Aspose.Cells that iterates over Workbook.Worksheets.ExternalLinks and replaces a given domain in OriginalDataSource. | Show how to assign a custom RibbonXml string to a Workbook object to create a new ribbon tab, then save as .xlsm. | Explain how to verify that external link URLs were updated and that RibbonXml is present after modifications.

using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkAndRibbonDemo
{
    // C# example that creates a workbook, replaces old SharePoint URLs in all ExternalLinks, injects custom Ribbon XML to add a new tab, and saves the file as a macro‑enabled workbook using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (creation rule)
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 2. Add an external link to the workbook
            // -------------------------------------------------
            // The external file "SourceData.xlsx" contains sheets "Sheet1" and "Sheet2"
            int linkIndex = workbook.Worksheets.ExternalLinks.Add(
                "https://oldsharepoint.com/Projects/SourceData.xlsx",
                new string[] { "Sheet1", "Sheet2" });

            // -------------------------------------------------
            // 3. Change the external link path(s)
            // -------------------------------------------------
            // Iterate through all external links and replace the old base URL with a new one
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];

                // OriginalDataSource holds the stored path; modify it as needed
                string originalPath = link.OriginalDataSource;
                string updatedPath = originalPath.Replace(
                    "https://oldsharepoint.com/Projects/",
                    "https://newsharepoint.com/Shared/");

                // Apply the modified path back to the link
                link.OriginalDataSource = updatedPath;
            }

            // -------------------------------------------------
            // 4. Customize the Ribbon UI using RibbonXml property
            // -------------------------------------------------
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Custom Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the custom Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // -------------------------------------------------
            // 5. Save the workbook (save rule)
            // -------------------------------------------------
            // The workbook is saved as a macro-enabled file to preserve the Ribbon customization
            workbook.Save("ExternalLinkAndRibbonDemo.xlsm");

            // -------------------------------------------------
            // 6. Verify changes (optional console output)
            // -------------------------------------------------
            Console.WriteLine("External link paths after modification:");
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                Console.WriteLine($" - {link.OriginalDataSource}");
            }

            Console.WriteLine($"Ribbon XML set: {(string.IsNullOrEmpty(workbook.RibbonXml) ? "No" : "Yes")}");
            Console.WriteLine("Workbook saved as 'ExternalLinkAndRibbonDemo.xlsm'.");
        }
    }
}
