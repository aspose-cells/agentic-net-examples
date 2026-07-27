using System;
using System.IO;
using Aspose.Cells;

class RibbonXmlImportDemo
{
    static void Main()
    {
        // Path to the existing workbook (must be macro-enabled to hold ribbon UI)
        string workbookPath = "input.xlsm";

        // Load the workbook from file
        Workbook workbook = new Workbook(workbookPath);

        // Path to the external ribbon XML file
        string ribbonXmlPath = "customRibbon.xml";

        // Read the entire XML content from the file
        string ribbonXml = File.ReadAllText(ribbonXmlPath);

        // Assign the XML to the workbook's RibbonXml property
        workbook.RibbonXml = ribbonXml;

        // Save the workbook with the updated ribbon UI
        workbook.Save("output_with_ribbon.xlsm");

        // Simple verification output
        Console.WriteLine("Ribbon XML imported successfully: " + (workbook.RibbonXml != null));
    }
}