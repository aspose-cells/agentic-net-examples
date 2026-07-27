using System;
using System.IO;
using Aspose.Cells;

namespace ExternalLinkAndRibbonDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Add an external link to the workbook
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                int linkIndex = externalLinks.Add(
                    "https://oldserver.com/Folder/Source.xlsx",
                    new[] { "Sheet1!A1" });

                ExternalLink link = externalLinks[linkIndex];

                // Verify the original data source
                if (link.OriginalDataSource != "https://oldserver.com/Folder/Source.xlsx")
                    throw new InvalidOperationException("Original data source mismatch.");

                // Modify the external link path
                string modifiedPath = link.OriginalDataSource.Replace(
                    "https://oldserver.com/Folder/",
                    "/newpath/");

                link.OriginalDataSource = modifiedPath;

                // Verify the path was changed correctly
                if (link.OriginalDataSource != "/newpath/Source.xlsx")
                    throw new InvalidOperationException("Modified path mismatch.");

                // Set custom Ribbon XML
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

                // Verify RibbonXml property is set
                if (string.IsNullOrEmpty(workbook.RibbonXml) || workbook.RibbonXml != ribbonXml)
                    throw new InvalidOperationException("Ribbon XML not set correctly.");

                // Save the workbook to a memory stream (XLSM to preserve Ribbon XML)
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsm);
                    ms.Position = 0;

                    // Load the workbook from the stream
                    var loadedWorkbook = new Workbook(ms);

                    // Verify Ribbon XML persisted after reload
                    if (loadedWorkbook.RibbonXml != ribbonXml)
                        throw new InvalidOperationException("Ribbon XML did not persist after reload.");

                    // Verify external link path persisted after reload
                    ExternalLink loadedLink = loadedWorkbook.Worksheets.ExternalLinks[0];
                    if (loadedLink.OriginalDataSource != "/newpath/Source.xlsx")
                        throw new InvalidOperationException("External link path did not persist after reload.");
                }

                Console.WriteLine("All validations passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}