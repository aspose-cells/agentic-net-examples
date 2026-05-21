using System;
using System.IO;
using Aspose.Cells;

namespace CustomRibbonExternalLinkDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Add a sample worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Insert a formula that references an external workbook (creates an external link)
                // Correct external link syntax: ='[full_path]SheetName'!CellReference
                sheet.Cells["A1"].Formula = "='[https://oldserver.com/Shared/OldData.xlsx]Sheet1'!A1";

                // ----- Update external link paths -----
                // Iterate through all external links and replace the old base URL with the new one
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    // Use OriginalDataSource to keep the original value and modify it
                    string original = externalLinks[i].OriginalDataSource;
                    string updated = original.Replace(
                        "https://oldserver.com/Shared/",
                        "https://newserver.com/Resources/");

                    externalLinks[i].OriginalDataSource = updated;
                }

                // ----- Set custom Ribbon UI -----
                // Define the Ribbon XML that will be embedded in the workbook
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab id=\"customTab\" label=\"My Custom Tab\">" +
                    "        <group id=\"customGroup\" label=\"My Group\">" +
                    "          <button id=\"customButton\" label=\"Refresh Links\" size=\"large\" onAction=\"RefreshLinks\" />" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";

                // Assign the Ribbon XML to the workbook (property: RibbonXml)
                workbook.RibbonXml = ribbonXml;

                // Optionally, prevent Excel from prompting to update links when the file opens
                workbook.Settings.UpdateLinksType = UpdateLinksType.Never;

                // Define output file path
                string outputPath = "CustomRibbonWithUpdatedLinks.xlsm";

                // Save the workbook (lifecycle: save)
                workbook.Save(outputPath);

                // Output confirmation
                Console.WriteLine($"Workbook saved with updated external links and custom Ribbon UI at '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}