using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportRibbonXmlDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define custom ribbon XML
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

                // Assign the RibbonXml to the workbook
                workbook.RibbonXml = ribbonXml;

                // Save the workbook as a macro‑enabled file (required lifecycle step)
                string workbookPath = "WorkbookWithRibbon.xlsm";
                workbook.Save(workbookPath, SaveFormat.Xlsm);

                // Export the RibbonXml to a separate .xml file for version control
                string exportPath = "RibbonCustom.xml";
                File.WriteAllText(exportPath, workbook.RibbonXml);

                Console.WriteLine($"Ribbon XML exported to {exportPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportRibbonXmlDemo.Run();
        }
    }
}