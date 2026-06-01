using System;
using Aspose.Cells;

namespace AsposeCellsDeveloperGuide
{
    public class RibbonAndExternalLinkDemo
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Define custom Ribbon XML
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

                // 4. Add an external link (formula referencing an external file)
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].Formula = "='https://example.com/OldFolder/ExternalFile.xlsx'!Sheet1!A1";

                // 5. Modify the external link paths using OriginalDataSource
                for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
                {
                    string originalLink = workbook.Worksheets.ExternalLinks[i].OriginalDataSource;
                    string modifiedLink = originalLink.Replace(
                        "https://example.com/OldFolder/",
                        "https://example.com/NewFolder/");

                    workbook.Worksheets.ExternalLinks[i].OriginalDataSource = modifiedLink;
                }

                // 6. Save the workbook as a macro‑enabled file to preserve the Ribbon XML
                workbook.Save("RibbonAndExternalLinkDemo.xlsm");

                // 7. Verification output
                Console.WriteLine("Ribbon XML set: " + (workbook.RibbonXml != null));
                foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
                {
                    Console.WriteLine("Modified External Link: " + link.OriginalDataSource);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
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