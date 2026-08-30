// Title: Import a custom Ribbon XML file into a macro-enabled Excel workbook with Aspose.Cells for .NET
// AI Prompts: Read the contents of a Ribbon XML file, assign it to the Workbook.RibbonXml property, and save the workbook. | Load an existing .xlsm workbook using Aspose.Cells, replace its Ribbon definition with external XML, and write the updated file. | Programmatically customize the Ribbon UI of a macro-enabled workbook by setting RibbonXml from a file path in C#.
// Common Searches: c# Aspose.Cells how to set RibbonXml for an .xlsm workbook | import external ribbon xml into existing Excel file using Aspose.Cells | replace ribbon UI in macro-enabled workbook programmatically | example of loading custom ribbon definition with Aspose.Cells .NET | update Excel Ribbon from XML file using C# Aspose.Cells
// Tags: RibbonXml property Aspose.Cells | set RibbonXml from external file C# | macro-enabled workbook ribbon customization | update ribbon UI Aspose.Cells .NET | load custom ribbon definition into xlsm

using System;
using System.IO;
using Aspose.Cells;

// The example loads a macro-enabled workbook, reads a custom Ribbon XML file, assigns the XML to the workbook's RibbonXml property, and saves the workbook with the updated Ribbon UI.
class RibbonXmlImportDemo
{
    static void Main()
    {
        // Path to the existing workbook (must be macro-enabled to support Ribbon UI)
        string workbookPath = "input.xlsm";

        // Path to the external Ribbon XML file
        string ribbonXmlPath = "customRibbon.xml";

        // Load the existing workbook
        Workbook workbook = new Workbook(workbookPath);

        // Read the Ribbon XML content from the external file
        string ribbonXml = File.ReadAllText(ribbonXmlPath);

        // Assign the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook with the updated Ribbon UI
        workbook.Save("output.xlsm");
    }
}
