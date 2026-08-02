using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define Ribbon XML that adds a custom tab named "Data Tools"
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"dataToolsTab\" label=\"Data Tools\">" +
            "        <group id=\"dataToolsGroup\" label=\"Data Tools Group\">" +
            "          <button id=\"sampleButton\" label=\"Sample\" size=\"large\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (as a macro-enabled file to preserve the ribbon UI)
        workbook.Save("DataToolsRibbon.xlsm");
    }
}