using System;
using Aspose.Cells;

class RibbonExternalLinkDemo
{
    static void Main()
    {
        // Load the workbook that contains external links
        Workbook workbook = new Workbook("input.xlsx");

        // Update each external link's path to the new location
        for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
        {
            ExternalLink link = workbook.Worksheets.ExternalLinks[i];
            string originalPath = link.OriginalDataSource;

            // Example: replace the old SharePoint base URL with the new one
            string updatedPath = originalPath.Replace(
                @"https://oldsharepoint.com/Fund II/",
                @"/sites/shared/shared documents/Fund II/");

            link.OriginalDataSource = updatedPath;
        }

        // Define custom Ribbon UI XML
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

        // Assign the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Prevent Excel from prompting to update external links when the file is opened
        workbook.Settings.UpdateLinksType = UpdateLinksType.Never;

        // Save the workbook (use .xlsm to retain the custom Ribbon UI)
        workbook.Save("output.xlsm");
    }
}