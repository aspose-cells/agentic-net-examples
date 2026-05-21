using System;
using System.IO;
using Aspose.Cells;

namespace RibbonXmlImportDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that will receive the custom Ribbon UI
            string workbookPath = "input.xlsx";

            // Path to the external XML file that defines the Ribbon UI
            string ribbonXmlPath = "customRibbon.xml";

            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Read the Ribbon XML content from the external file
            string ribbonXml = File.ReadAllText(ribbonXmlPath);

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook (using .xlsm to preserve the Ribbon UI)
            workbook.Save("output.xlsm");

            Console.WriteLine("Custom Ribbon XML has been imported and workbook saved as 'output.xlsm'.");
        }
    }
}