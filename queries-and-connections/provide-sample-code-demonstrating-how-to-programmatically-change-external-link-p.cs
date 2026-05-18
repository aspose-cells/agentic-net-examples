using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExternalLinkAndRibbonDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Customize Ribbon --------------------
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

            workbook.RibbonXml = ribbonXml;

            // -------------------- Add an external link (for demo) --------------------
            workbook.Worksheets.ExternalLinks.Add(
                "https://oldserver.com/SharedDocs/Finance.xlsx",
                new string[] { "Sheet1!A1", "Sheet1!B2" });

            // -------------------- Change external link paths --------------------
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                string originalPath = workbook.Worksheets.ExternalLinks[i].OriginalDataSource;
                string newBase = @"https://newserver.com/Docs/Finance/";
                string modifiedPath = originalPath.Replace(
                    "https://oldserver.com/SharedDocs/",
                    newBase);
                workbook.Worksheets.ExternalLinks[i].OriginalDataSource = modifiedPath;
            }

            // Verify changes
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                Console.WriteLine("Modified External Link: " + link.OriginalDataSource);
            }

            // Save the workbook
            workbook.Save("ExternalLinkAndRibbonDemo.xlsm");

            Console.WriteLine("Workbook saved with customized ribbon and updated external links.");
        }
    }
}