using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define Ribbon XML with a custom tab, group, and a button.
        // The button calls a macro named "RefreshLinks" when clicked.
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"External Links\">" +
            "        <group id=\"linkGroup\" label=\"Operations\">" +
            "          <button id=\"refreshLinksBtn\" label=\"Refresh Links\" size=\"large\" onAction=\"RefreshLinks\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook as a macro-enabled file (xlsm) so the macro can be added later.
        workbook.Save("WorkbookWithCustomRibbon.xlsm");
    }
}