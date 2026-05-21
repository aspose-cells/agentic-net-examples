using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExternalLinkAndRibbonDemo
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
            // -------------------------------------------------
            // 1. Create a new workbook (creation rule)
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 2. Add an external link to demonstrate path change
            // -------------------------------------------------
            // Add a link to a fictitious external file with a sample sheet name
            int linkIndex = workbook.Worksheets.ExternalLinks.Add(
                "https://oldserver.sharepoint.com/Shared%20Documents/Finance/Report.xlsx",
                new string[] { "Sheet1" });

            // Verify the added link
            ExternalLink link = workbook.Worksheets.ExternalLinks[linkIndex];
            Console.WriteLine("OriginalDataSource before change: " + link.OriginalDataSource);

            // -------------------------------------------------
            // 3. Change the external link path (using OriginalDataSource)
            // -------------------------------------------------
            // Replace the old server URL with a new one
            string modifiedPath = link.OriginalDataSource.Replace(
                "https://oldserver.sharepoint.com/Shared%20Documents/",
                "https://newserver.sharepoint.com/Docs/Finance/");

            // Apply the modified path back to the external link
            link.OriginalDataSource = modifiedPath;

            Console.WriteLine("OriginalDataSource after change: " + link.OriginalDataSource);

            // -------------------------------------------------
            // 4. Customize the Ribbon UI (using RibbonXml property)
            // -------------------------------------------------
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Custom Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the custom Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // -------------------------------------------------
            // 5. Save the workbook (save rule)
            // -------------------------------------------------
            string outputPath = "ExternalLinkAndRibbonDemo.xlsm";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved with updated external link and custom Ribbon.");
        }
    }
}