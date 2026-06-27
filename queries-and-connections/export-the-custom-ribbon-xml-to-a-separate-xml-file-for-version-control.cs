using System;
using System.IO;
using Aspose.Cells;

class ExportRibbonXml
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Define the custom Ribbon XML
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

        // Export the Ribbon XML to a separate .xml file for version control
        string exportPath = "RibbonCustom.xml";
        File.WriteAllText(exportPath, workbook.RibbonXml);

        Console.WriteLine($"Ribbon XML has been exported to '{exportPath}'.");
    }
}