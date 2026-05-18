using System;
using System.IO;
using Aspose.Cells;

class ExportRibbonXml
{
    static void Main()
    {
        // Create a new workbook (uses the provided create rule)
        Workbook workbook = new Workbook();

        // Sample custom ribbon XML
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

        // Set the RibbonXml property (provided member)
        workbook.RibbonXml = ribbonXml;

        // Export the RibbonXml to a separate .xml file for version control
        string ribbonFilePath = "CustomRibbon.xml";
        File.WriteAllText(ribbonFilePath, workbook.RibbonXml);

        // Save the workbook itself (uses the provided save rule)
        workbook.Save("WorkbookWithRibbon.xlsm");

        Console.WriteLine($"Ribbon XML exported to '{ribbonFilePath}'.");
    }
}