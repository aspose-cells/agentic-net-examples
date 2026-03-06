using System;
using Aspose.Cells;

namespace RibbonXmlPreserveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default XLSX format)
            Workbook workbook = new Workbook();

            // Define the custom Ribbon XML you want to keep unchanged
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

            // Assign the Ribbon XML to the workbook; it will be stored without parsing
            workbook.RibbonXml = ribbonXml;

            // Save the workbook. Using XLSM ensures the Ribbon XML is retained in the package.
            workbook.Save("PreservedRibbon.xlsm");

            // Verify that the RibbonXml property is still set after saving (optional)
            Console.WriteLine("RibbonXml preserved: " + (workbook.RibbonXml != null));
        }
    }
}