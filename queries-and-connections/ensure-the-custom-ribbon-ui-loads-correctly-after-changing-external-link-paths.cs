// Title: Update external link URL in an Aspose.Cells workbook while preserving custom Ribbon UI and disabling update prompts (C#)
// AI Prompts: Generate C# code that assigns custom Ribbon XML to a workbook, adds an external link, replaces its OriginalDataSource and DataSource with a new URL, sets UpdateLinksType to Never, and saves the file as a macro‑enabled .xlsm. | Show how to modify the path of an existing ExternalLink in Aspose.Cells without losing the attached Ribbon customization. | Create a snippet that prevents Excel from displaying the external links update dialog after changing link URLs using Aspose.Cells.
// Common Searches: aspocells change external link path while keeping custom ribbon UI | c# update external workbook link URL in Aspose.Cells and suppress update prompt | save workbook with ribbon XML after modifying external link in Aspose.Cells | how to replace OriginalDataSource of an ExternalLink in Aspose.Cells
// Tags: update external link URL Aspose.Cells | preserve custom ribbon XML Aspose.Cells | disable external links update prompt Aspose.Cells | save macro-enabled workbook with ribbon Aspose.Cells | modify OriginalDataSource ExternalLink C#

using System;
using Aspose.Cells;

namespace RibbonAndExternalLinkDemo
{
    // The example creates a new workbook, embeds custom Ribbon XML, adds an external link, updates the link's URL by replacing the old path in OriginalDataSource and DataSource, disables automatic link updates, and saves the result as a macro‑enabled .xlsm file, ensuring the Ribbon UI remains functional.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // 1. Define custom Ribbon UI XML and assign it to the workbook
            // -----------------------------------------------------------------
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

            // Set the RibbonXml property so Excel will load the custom UI
            workbook.RibbonXml = ribbonXml;

            // -----------------------------------------------------------------
            // 2. Add an external link (simulating a link that will later be changed)
            // -----------------------------------------------------------------
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            // Add a dummy external workbook reference
            int linkIndex = externalLinks.Add("https://oldserver.com/Shared/OldFolder/ExternalData.xlsx",
                                             new string[] { "Sheet1" });

            // -----------------------------------------------------------------
            // 3. Update the external link path to the new location
            // -----------------------------------------------------------------
            ExternalLink link = externalLinks[linkIndex];
            // Use OriginalDataSource to keep the original value if needed; here we replace it
            string updatedPath = link.OriginalDataSource.Replace(
                                    "https://oldserver.com/Shared/OldFolder/",
                                    "https://newserver.com/Shared/NewFolder/");
            link.OriginalDataSource = updatedPath;

            // Optionally also update DataSource (the active path used by formulas)
            link.DataSource = updatedPath;

            // -----------------------------------------------------------------
            // 4. Prevent Excel from prompting to update external links on open
            // -----------------------------------------------------------------
            workbook.Settings.UpdateLinksType = UpdateLinksType.Never;

            // -----------------------------------------------------------------
            // 5. Save the workbook (as a macro-enabled file to retain Ribbon XML)
            // -----------------------------------------------------------------
            workbook.Save("RibbonWithUpdatedLinks.xlsm");

            // Confirmation output
            Console.WriteLine("Workbook saved with custom Ribbon UI and updated external link path.");
        }
    }
}
