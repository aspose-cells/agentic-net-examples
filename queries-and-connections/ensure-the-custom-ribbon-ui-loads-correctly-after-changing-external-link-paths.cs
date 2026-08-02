using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RibbonAndExternalLinkDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsm";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException($"Input file not found: {inputPath}");
            }

            // Load an existing workbook that contains external links
            Workbook workbook = new Workbook(inputPath);

            // ------------------------------------------------------------
            // 1. Update external link paths (e.g., change SharePoint URL)
            // ------------------------------------------------------------
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                // Keep the original data source path
                string originalPath = externalLinks[i].OriginalDataSource;

                // Example replacement: change old SharePoint base URL to new one
                string updatedPath = originalPath.Replace(
                    @"https://oldsharepoint.com/Docs/",
                    @"/sites/newsite/shared documents/");

                // Apply the modified path
                externalLinks[i].OriginalDataSource = updatedPath;
            }

            // ------------------------------------------------------------
            // 2. Define custom Ribbon UI XML
            // ------------------------------------------------------------
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Custom Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"Refresh Links\" size=\"large\" " +
                "                  onAction=\"OnRefreshLinks\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // ------------------------------------------------------------
            // 3. Prevent Excel from prompting to update external links on open
            // ------------------------------------------------------------
            workbook.Settings.UpdateLinksType = UpdateLinksType.Never;

            // ------------------------------------------------------------
            // 4. Save the workbook (use .xlsm to retain Ribbon UI)
            // ------------------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}