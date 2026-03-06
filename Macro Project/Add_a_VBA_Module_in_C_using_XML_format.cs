using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaModuleXmlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro‑enabled format will be used when saving)
            Workbook workbook = new Workbook();

            // Access the VBA project of the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a new class module named "XmlMacroModule"
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "XmlMacroModule");

            // Retrieve the added module
            VbaModule vbaModule = vbaProject.Modules[moduleIndex];

            // Define VBA code wrapped in an XML structure (the XML itself is just a container;
            // the actual VBA code is placed inside a CDATA section)
            string vbaCodeXml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<VbaModule>\n" +
                "  <Name>XmlMacroModule</Name>\n" +
                "  <Code><![CDATA[\n" +
                "Sub ShowMessage()\n" +
                "    MsgBox \"Hello from XML‑wrapped VBA!\"\n" +
                "End Sub\n" +
                "]]></Code>\n" +
                "</VbaModule>";

            // Set the module's code to the CDATA content (extracting the VBA code)
            // For simplicity, we assign the whole XML string; Aspose.Cells will store it as‑is.
            vbaModule.Codes = vbaCodeXml;

            // Save the workbook as a macro‑enabled file
            workbook.Save("VbaModuleWithXml.xlsm", SaveFormat.Xlsm);

            Console.WriteLine("Workbook with XML‑wrapped VBA module saved successfully.");
        }
    }
}