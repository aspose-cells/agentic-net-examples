using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RibbonAndExternalLinkDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define custom Ribbon XML
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

            // Assign Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Add a sample external link (simulated)
            string sharepointUrl = "https://arcusventures.sharepoint.com/Fund II/example.xlsx";
            workbook.Worksheets[0].Cells["A1"].Formula = $"='[{sharepointUrl}]Sheet1'!A1";

            // Modify external link paths
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];
                string originalPath = link.OriginalDataSource;
                string updatedPath = originalPath.Replace(
                    @"https://arcusventures.sharepoint.com/Fund II/",
                    @"/sites/shared/shared documents/Fund II/");
                link.OriginalDataSource = updatedPath;
            }

            // Define output file
            string outputPath = "RibbonAndExternalLinkDemo.xlsm";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}