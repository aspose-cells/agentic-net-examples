// Title: C# – Update External Link URLs and Embed a Custom Ribbon UI in an Excel .xlsm Workbook with Aspose.Cells
// Description: Loads an .xlsm file, replaces an old SharePoint base path in every external link, embeds a custom Ribbon XML definition, disables automatic link updates, and saves the workbook. Demonstrates how to keep formulas intact while modernizing data sources and adding a personalized Ribbon tab.
// Keywords: Aspose.Cells C# external links | replace SharePoint URL Excel | custom Ribbon XML Aspose.Cells | prevent Excel update links prompt | modify workbook ribbon programmatically | UpdateLinksType.Never | Excel .xlsm custom tab | C# Excel automation Aspose
// Common Searches: replace old SharePoint URL in Excel external links using Aspose.Cells | add custom Ribbon tab to .xlsm workbook with C# | disable link update dialog when opening Excel file via Aspose.Cells | update external link paths after SharePoint migration C# | embed Ribbon XML in workbook programmatically
// Developer Intent: Refresh external link paths, embed a custom Ribbon UI, and suppress link‑update prompts before saving the workbook.
// Use Cases: Migrate Excel workbooks to a new SharePoint site without breaking external data connections. | Provide users with a dedicated Ribbon tab that launches macros or custom actions. | Open workbooks in Excel without prompting users to refresh external connections.
// AI Prompts: Write C# code with Aspose.Cells that replaces a base URL in all external links of a workbook and adds custom Ribbon XML. | Show how to set UpdateLinksType to Never in Aspose.Cells to stop Excel from asking to refresh external links. | Explain how to verify that the custom Ribbon UI is correctly embedded in the saved .xlsm file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an .xlsm file, replaces an old SharePoint base path in every external link, embeds a custom Ribbon XML definition, disables automatic link updates, and saves the workbook. Demonstrates how to keep formulas intact while modernizing data sources and adding a personalized Ribbon tab.
    public class RibbonAndExternalLinkDemo
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsm";
                const string outputPath = "output.xlsm";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load an existing workbook that contains external links
                Workbook workbook = new Workbook(inputPath);

                // Define the old and new base paths for external links
                string oldBasePath = @"https://oldsharepoint.com/Projects/";
                string newBasePath = @"/sites/shared/Projects/";

                // Update each external link's original data source to the new path
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    string original = externalLinks[i].OriginalDataSource;
                    if (!string.IsNullOrEmpty(original) && original.Contains(oldBasePath))
                    {
                        string updated = original.Replace(oldBasePath, newBasePath);
                        externalLinks[i].OriginalDataSource = updated;
                    }
                }

                // Sample custom Ribbon XML to be embedded in the workbook
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

                // Assign the Ribbon XML to the workbook
                workbook.RibbonXml = ribbonXml;

                // Prevent Excel from prompting to update external links when the file is opened
                workbook.Settings.UpdateLinksType = UpdateLinksType.Never;

                // Save the modified workbook; the custom ribbon UI will be available on open
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
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
