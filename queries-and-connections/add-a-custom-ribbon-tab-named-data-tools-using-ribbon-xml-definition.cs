using System;
using Aspose.Cells;

namespace AsposeCellsRibbonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Define the Ribbon XML that adds a custom tab named "Data Tools"
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"dataToolsTab\" label=\"Data Tools\">" +
                "        <group id=\"dataToolsGroup\" label=\"Data Operations\">" +
                "          <button id=\"refreshButton\" label=\"Refresh\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro-enabled file (XLSM) to preserve the Ribbon UI
            workbook.Save("DataToolsDemo.xlsm", SaveFormat.Xlsm);

            // Optional: confirm that the Ribbon XML was set
            Console.WriteLine("Custom Ribbon tab 'Data Tools' has been added and workbook saved.");
        }
    }
}