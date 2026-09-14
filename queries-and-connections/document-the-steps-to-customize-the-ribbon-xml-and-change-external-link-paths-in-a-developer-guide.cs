// Title: How to customize the Excel ribbon with XML and rewrite external SharePoint link paths using Aspose.Cells for .NET
// AI Prompts: Generate C# code that assigns a custom Ribbon XML string to Workbook.RibbonXml, adds a SharePoint external link, replaces its base URL with a relative path, and saves the workbook as a macro‑enabled .xlsm file using Aspose.Cells. | Demonstrate how to iterate over Workbook.Worksheets.ExternalLinks and update each link's OriginalDataSource value in Aspose.Cells. | Provide a step‑by‑step guide for preserving a custom ribbon UI when saving a macro‑enabled workbook with Aspose.Cells.
// Common Searches: aspnet add custom tab to Excel ribbon using Aspose.Cells | replace base SharePoint URL in external links of an Aspose.Cells workbook | set RibbonXml property and save as macro enabled workbook in C# | update external link paths programmatically with Aspose.Cells .NET | preserve custom ribbon UI when saving .xlsm with Aspose.Cells
// Tags: Workbook.RibbonXml customization | custom ribbon XML Aspose.Cells | external link URL rewrite .NET | modify ExternalLinkCollection paths | save macro-enabled xlsm Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsDeveloperGuide
{
    // This example shows how to create a new Workbook, define custom ribbon XML, assign it via the RibbonXml property, add a SharePoint external link, loop through the ExternalLinkCollection to replace the SharePoint base URL with a relative path, and finally save the workbook as a macro‑enabled .xlsm file while outputting verification details.
    public class RibbonAndExternalLinkDemo
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // 2. Define custom Ribbon XML
                //    This XML adds a new tab with a group and a button to the Excel ribbon.
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab id=\"customTab\" label=\"My Tab\">" +
                    "        <group id=\"customGroup\" label=\"My Group\">" +
                    "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";

                // 3. Assign the Ribbon XML to the workbook (property: RibbonXml)
                workbook.RibbonXml = ribbonXml;

                // 4. Add an external link to demonstrate path modification
                //    The link points to a sample file on a SharePoint site.
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                string originalPath = @"https://arcusventures.sharepoint.com/Fund II/example.xlsx";
                externalLinks.Add(originalPath, new string[] { "Sheet1" });

                // 5. Iterate through all external links and replace the SharePoint URL
                //    with a relative path used in the target environment.
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    // Get the stored original data source (property: OriginalDataSource)
                    string currentLink = externalLinks[i].OriginalDataSource;

                    // Replace the base URL with the new relative path
                    string modifiedLink = currentLink.Replace(
                        @"https://arcusventures.sharepoint.com/Fund II/",
                        @"/sites/shared/shared documents/Fund II/");

                    // Update the external link with the new path
                    externalLinks[i].OriginalDataSource = modifiedLink;
                }

                // 6. Save the workbook (lifecycle: save)
                //    The workbook is saved as an Excel macro-enabled file to preserve the Ribbon UI.
                workbook.Save("CustomizedRibbonAndLinks.xlsm");

                // 7. Optional verification output
                Console.WriteLine("Ribbon XML set: " + (workbook.RibbonXml != null));
                Console.WriteLine("External links after modification:");
                foreach (ExternalLink link in externalLinks)
                {
                    Console.WriteLine("- " + link.OriginalDataSource);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RibbonAndExternalLinkDemo.Run();
        }
    }
}
