using System;
using Aspose.Cells;

namespace ExternalLinkRibbonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add an external link formula (valid external reference syntax)
            sheet.Cells["A1"].Formula = "='[SourceData.xlsx]Sheet1'!A1";

            // Define custom Ribbon XML with a button (the button would be wired to a macro in a real scenario)
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"External Link Tools\">" +
                "        <group id=\"linkGroup\" label=\"Link Operations\">" +
                "          <button id=\"updateLinkButton\" label=\"Update Links\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Iterate through all external links and update their OriginalDataSource paths
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];
                string original = link.OriginalDataSource;

                // Example transformation: replace old server URL with new server URL
                string updated = original.Replace(
                    "https://oldserver.com/Folder/",
                    "https://newserver.com/Shared/");

                // Apply the updated path
                link.OriginalDataSource = updated;
            }

            // Save the workbook (as a macro-enabled file to retain Ribbon XML)
            workbook.Save("ExternalLinkUpdated.xlsm");

            // Output verification information
            Console.WriteLine("External links after update:");
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                Console.WriteLine($"OriginalDataSource: {link.OriginalDataSource}");
            }
        }
    }
}