using System;
using System.IO;
using Aspose.Cells;

namespace RibbonXmlImportDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook (must be macro-enabled to support Ribbon UI)
            string workbookPath = "Template.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Path to the external Ribbon XML file
            string ribbonXmlPath = "CustomRibbon.xml";

            // Read the entire XML content
            string ribbonXml = File.ReadAllText(ribbonXmlPath);

            // Assign the XML to the workbook's RibbonXml property
            workbook.RibbonXml = ribbonXml;

            // Save the workbook with the updated Ribbon UI
            string outputPath = "WorkbookWithCustomRibbon.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Custom Ribbon XML applied and saved to '{outputPath}'.");
        }
    }
}