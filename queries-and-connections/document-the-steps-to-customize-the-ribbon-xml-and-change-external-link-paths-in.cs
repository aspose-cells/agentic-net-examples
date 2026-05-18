using System;
using Aspose.Cells;

namespace AsposeCellsDeveloperGuide
{
    public class RibbonAndExternalLinkDemo
    {
        public static void Run()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // 2. Define custom Ribbon XML
            //    This XML will replace the default Ribbon UI when the workbook is opened in Excel.
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

            // 4. Add a sample external link so we have something to modify
            //    Use a valid external reference formula.
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].Formula = "='[Source.xlsx]Sheet1'!A1";

            // 5. Iterate through all external links and change their stored source path
            //    Here we replace the old SharePoint URL with a new relative path.
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                // Get the original stored path
                string originalPath = externalLinks[i].OriginalDataSource;

                // Replace the old base URL with the new one
                string modifiedPath = originalPath.Replace(
                    "https://example.sharepoint.com/Docs/",
                    "/shared/documents/");

                // Update the external link with the new path
                externalLinks[i].OriginalDataSource = modifiedPath;
            }

            // 6. Save the workbook (lifecycle: save)
            //    The file is saved as a macro‑enabled workbook to preserve the Ribbon XML.
            workbook.Save("CustomizedRibbonAndLinks.xlsm");

            // 7. Simple verification output
            Console.WriteLine("Ribbon XML set: " + (workbook.RibbonXml != null));
            Console.WriteLine("External links after modification:");
            foreach (ExternalLink link in externalLinks)
            {
                Console.WriteLine("- " + link.OriginalDataSource);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RibbonAndExternalLinkDemo.Run();
        }
    }
}