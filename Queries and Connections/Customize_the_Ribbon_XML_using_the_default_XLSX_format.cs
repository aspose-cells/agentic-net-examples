using System;
using Aspose.Cells;

namespace AsposeCellsRibbonDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (default XLSX format)
            Workbook workbook = new Workbook();

            // Define custom Ribbon XML
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

            // Create OOXML save options for XLSX format
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Save the workbook with the custom Ribbon UI
            workbook.Save("CustomRibbonWorkbook.xlsx", saveOptions);

            // Verify that RibbonXml is set (optional)
            Console.WriteLine("RibbonXml set: " + (workbook.RibbonXml != null));
        }
    }
}