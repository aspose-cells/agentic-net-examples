// Title: Aspose.Cells .NET – Customize Excel Ribbon with XML and Rewrite External Link URLs
// Description: Step‑by‑step guide that creates a new Workbook, injects custom Ribbon XML to add a tab, group and button, inserts a formula that points to an external workbook, iterates the ExternalLinks collection to replace an old base URL with a new SharePoint path, and saves the file as a macro‑enabled .xlsm to retain the custom UI.
// Keywords: Aspose.Cells | custom ribbon XML | Workbook.RibbonXml | external link URL update | ExternalLink.OriginalDataSource | sharepoint path replacement | xlsm macro enabled | C# .NET example | Excel UI customization | global
// Common Searches: how to add a custom tab to Excel ribbon using Aspose.Cells | replace external workbook links in Aspose.Cells workbook | save workbook with custom ribbon as xlsm | update SharePoint URLs in Excel external links via C# | Aspose.Cells RibbonXml property example
// Developer Intent: Generate a workbook that displays a custom ribbon UI and contains updated external link URLs before persisting it as a macro‑enabled file.
// Use Cases: Inject custom Ribbon XML to introduce new UI elements (tab, group, button) in Excel. | Programmatically modify every external link's source path to point to a new SharePoint location. | Persist the custom ribbon definition by saving the workbook as an .xlsm file.
// AI Prompts: Write C# code that sets Workbook.RibbonXml with custom UI markup and saves the workbook as .xlsm using Aspose.Cells. | Show how to loop through Workbook.Worksheets.ExternalLinks and replace each link's OriginalDataSource with a new SharePoint URL. | Explain how to verify that the custom ribbon appears and that external links resolve to the updated paths after the file is opened.

using System;
using Aspose.Cells;

namespace AsposeCellsDeveloperGuide
{
    // Step‑by‑step guide that creates a new Workbook, injects custom Ribbon XML to add a tab, group and button, inserts a formula that points to an external workbook, iterates the ExternalLinks collection to replace an old base URL with a new SharePoint path, and saves the file as a macro‑enabled .xlsm to retain the custom UI.
    public class RibbonAndExternalLinkDemo
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook instance
                Workbook workbook = new Workbook();

                // 2. Define the custom Ribbon XML that will replace the default ribbon UI
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

                // 3. Assign the Ribbon XML to the workbook
                workbook.RibbonXml = ribbonXml;

                // 4. Add a sample formula that references an external workbook (to create an external link)
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].Formula = "='https://example.com/OldFolder/ExternalFile.xlsx'!Sheet1!A1";

                // 5. Iterate through all external links and modify their stored paths
                for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
                {
                    ExternalLink link = workbook.Worksheets.ExternalLinks[i];
                    string originalPath = link.OriginalDataSource;

                    // Replace the old base URL with the new location
                    string newPath = originalPath.Replace(
                        "https://example.com/OldFolder/",
                        "/sites/shared/documents/NewFolder/");

                    // Update the external link with the new path
                    link.OriginalDataSource = newPath;
                }

                // 6. Save the workbook as a macro‑enabled file (xlsm) to preserve the Ribbon XML
                workbook.Save("RibbonAndExternalLinkDemo.xlsm");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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
