using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExternalLinkAndRibbonDemo
    {
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a sample external link that we will later modify
            int linkIndex = workbook.Worksheets.ExternalLinks.Add(
                "https://oldserver.com/SharedDocs/Folder/Source.xlsx",
                new string[] { "Sheet1" });

            // Show the original external link path
            Console.WriteLine("Original External Link: " +
                workbook.Worksheets.ExternalLinks[linkIndex].OriginalDataSource);

            // Programmatically change external link paths
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];
                string original = link.OriginalDataSource;

                // Replace old server URL with a new one
                string modified = original.Replace(
                    "https://oldserver.com/SharedDocs/Folder/",
                    "https://newserver.com/Docs/");

                link.OriginalDataSource = modified;
                Console.WriteLine($"Modified Link [{i}]: {link.OriginalDataSource}");
            }

            // Custom Ribbon UI XML
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

            workbook.RibbonXml = ribbonXml;

            // Define output file path
            string outputPath = "ExternalLinkAndRibbonDemo.xlsm";

            // Save the workbook (macro-enabled format preserves Ribbon XML)
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}' with updated external links and custom Ribbon.");
        }
    }
}