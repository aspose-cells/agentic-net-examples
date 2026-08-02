using System;
using Aspose.Cells;

class RibbonExternalLinkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define Ribbon XML that adds a button.
        // The button calls a macro named "PromptExternalLink" when clicked.
        // The macro (to be added separately) should display a file‑open dialog,
        // let the user pick a workbook, and add it as an external link.
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"External Link\">" +
            "        <group id=\"linkGroup\" label=\"Link Operations\">" +
            "          <button id=\"promptButton\" label=\"Set External Link\" size=\"large\" onAction=\"PromptExternalLink\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the custom Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // When the workbook is opened, prompt the user to update external links
        workbook.Settings.UpdateLinksType = UpdateLinksType.UserSet;

        // Save the workbook as a macro‑enabled file (the macro itself must be added separately)
        workbook.Save("RibbonExternalLinkDemo.xlsm");

        Console.WriteLine("Workbook with custom Ribbon button created successfully.");
    }
}