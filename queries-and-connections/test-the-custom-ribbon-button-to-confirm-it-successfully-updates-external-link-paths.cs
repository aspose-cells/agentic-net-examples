using System;
using Aspose.Cells;

namespace RibbonExternalLinkUpdateDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Add an external link formula to cell A1 of the first worksheet
            //    The formula points to a sample SharePoint location
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "='https://sharepoint.example.com/OldFolder/[SampleData.xlsx]Sheet1'!A1";

            // 3. Define custom Ribbon XML with a button that would (in a real UI) trigger the link update
            //    Here we only set the XML; the button click is simulated programmatically below
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"Demo Tab\">" +
                "        <group id=\"linkGroup\" label=\"Link Operations\">" +
                "          <button id=\"updateLinksButton\" label=\"Update Links\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // 4. Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // 5. Simulate the custom ribbon button click:
            //    Iterate through all external links and replace the old base path with a new one
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                // Store the original data source for logging
                string original = link.OriginalDataSource;

                // Define the old and new base paths
                string oldBase = "https://sharepoint.example.com/OldFolder/";
                string newBase = "/sites/shared/shared documents/NewFolder/";

                // Perform the replacement
                string updated = original.Replace(oldBase, newBase);

                // Apply the updated path back to the external link
                link.OriginalDataSource = updated;

                // Optional: output the change to the console for verification
                Console.WriteLine($"Link updated from: {original}");
                Console.WriteLine($"Link updated to  : {updated}");
            }

            // 6. Save the workbook (including the updated external link and Ribbon XML)
            workbook.Save("RibbonExternalLinkUpdateDemo.xlsx");

            Console.WriteLine("Workbook saved. External link paths have been updated.");
        }
    }
}